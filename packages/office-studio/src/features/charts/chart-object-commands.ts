import type { OfficeStudio } from '../../api/office-studio'
import type { Disposable } from '../../api/events'
import type { ChartModel, WorkbookModel } from '../../domain/model'
import { canEditWorksheetProtection } from '../sheet-protection/range-protection'
import { assertChartReadPermission } from './services/chart-statistics-service'
import type { ChartEditPanelRoute } from './chart-edit-panel-route'

export const CHART_OBJECT_COMMANDS = {
  edit: 'sheet.command.chart.open-focused-panel',
  copy: 'sheet.command.chart.copy-focused-chart',
  exportImage: 'sheet.command.chart.export-focused-image',
  delete: 'sheet.command.chart.delete-focused-chart',
} as const
export const PASTE_COPIED_CHART_COMMAND = 'sheet.command.chart.paste-copied-chart'
export type ChartObjectAction = keyof typeof CHART_OBJECT_COMMANDS
export type ChartObjectCommandId = typeof CHART_OBJECT_COMMANDS[ChartObjectAction]
export const CHART_OBJECT_MENU: readonly { action: ChartObjectAction; label: string; order: number; shortcut?: string }[] = [
  { action: 'edit', label: '编辑图表', order: 0, shortcut: 'F2' },
  { action: 'copy', label: '复制图表', order: 1, shortcut: 'Ctrl+C' },
  { action: 'exportImage', label: '导出图片', order: 3 },
  { action: 'delete', label: '删除图表', order: 4, shortcut: 'Delete' },
]

/** Author identities and the optional UI session are control data, never model fields. */
export interface ChartObjectTarget {
  unitId: string
  sheetId: string
  chartId: string
  expectedRevision: number
  snapshot: string
  hostSessionId?: string
}
export interface ChartObjectCommandInput {
  unitId?: string
  sheetId?: string
  subUnitId?: string
  chartId?: string
  drawingId?: string
  expectedRevision?: number
  snapshot?: string
  hostSessionId?: string
  route?: ChartEditPanelRoute
}
export interface ChartObjectPasteInput {
  unitId?: string
  sheetId?: string
  subUnitId?: string
  expectedRevision?: number
  clipboardId?: string
  hostSessionId?: string
  newChartId?: string
  position?: { x: number; y: number }
}
export interface ChartObjectCapability { disabled: boolean; hidden?: boolean; reason?: string }
export type ChartObjectCapabilities = Record<ChartObjectAction, ChartObjectCapability>

/** Reserved, versioned chart-sheet adapter; ordinary grid sheets omit it.
 * Chart-sheet creation/import remains a separate host feature. */
export const CHART_SHEET_METADATA_KEY = 'officeChartSheet'
export function isPrimaryChartObject(workbook: WorkbookModel, chart: ChartModel): boolean {
  const descriptor = workbook.sheets.find(sheet => sheet.id === chart.sheetId)?.metadata?.[CHART_SHEET_METADATA_KEY]
  if (descriptor === undefined) return false
  if (!descriptor || typeof descriptor !== 'object' || Array.isArray(descriptor)
    || !('version' in descriptor) || descriptor.version !== 1 || !('primaryChartId' in descriptor)
    || typeof descriptor.primaryChartId !== 'string' || !descriptor.primaryChartId) throw new TypeError('图表工作表主图描述无效。')
  return descriptor.primaryChartId === chart.id
}

export function chartObjectCapabilities(workbook: WorkbookModel, chart: ChartModel, userId: string,
  options: { panel?: boolean; image?: boolean; busy?: boolean } = {}): ChartObjectCapabilities {
  let readReason: string | undefined, primary = false
  try { assertChartReadPermission(workbook, chart, userId); primary = isPrimaryChartObject(workbook, chart) }
  catch (cause) { readReason = cause instanceof Error ? cause.message : '图表或源数据不可查看。' }
  const sheet = workbook.sheets.find(sheet => sheet.id === chart.sheetId)
  const editReason = readReason ?? (!sheet || !canEditWorksheetProtection(sheet, userId) ? '工作表受保护，当前用户不能编辑图表。' : undefined)
  const entry = (reason?: string): ChartObjectCapability => ({ disabled: !!reason || !!options.busy,
    ...(reason ? { reason } : options.busy ? { reason: '图表操作正在处理。' } : {}) })
  return {
    edit: entry(editReason ?? (options.panel === false ? '图表属性面板宿主尚未绑定。' : undefined)),
    copy: entry(readReason),
    exportImage: entry(readReason ?? (options.image === false ? '图表绘制尚未完成。' : undefined)),
    delete: { ...entry(editReason ?? (primary ? '图表工作表的主图不能独立删除。' : undefined)), hidden: primary },
  }
}

export interface ChartObjectUIHost {
  sessionId: string
  target(): ChartObjectTarget | undefined
  current(target: ChartObjectTarget): boolean
  currentSheet(unitId: string, sheetId: string): boolean
  openPanel(target: ChartObjectTarget, route?: ChartEditPanelRoute): boolean
  exportImage(target: ChartObjectTarget): Promise<boolean>
  writeClipboard(content: { plain: string; html: string }): Promise<boolean>
}
const hosts = new WeakMap<OfficeStudio, ChartObjectUIHost[]>()
export function registerChartObjectUIHost(studio: OfficeStudio, host: ChartObjectUIHost): Disposable {
  const list = hosts.get(studio) ?? []
  if (list.some(entry => entry.sessionId === host.sessionId)) throw new Error('图表宿主会话重复。')
  list.push(host); hosts.set(studio, list)
  let active = true
  return { dispose() {
    if (!active) return
    active = false
    const current = hosts.get(studio)
    if (!current) return
    const index = current.indexOf(host)
    if (index >= 0) current.splice(index, 1)
    if (!current.length) hosts.delete(studio)
  } }
}
export function chartObjectUIHost(studio: OfficeStudio, target?: ChartObjectTarget): ChartObjectUIHost | undefined {
  return [...(hosts.get(studio) ?? [])].reverse().find(host => target
    ? (!target.hostSessionId || host.sessionId === target.hostSessionId) && host.current(target)
    : !!host.target())
}
export function chartObjectUISession(studio: OfficeStudio, sessionId: string): ChartObjectUIHost | undefined {
  return hosts.get(studio)?.find(host => host.sessionId === identifier(sessionId, '图表宿主会话'))
}
export function clearChartObjectUIHosts(studio: OfficeStudio): void { hosts.delete(studio) }

function identifier(value: unknown, name: string): string | undefined {
  if (value === undefined) return undefined
  if (typeof value !== 'string' || !value.trim() || value.length > 512) throw new TypeError(`${name} 无效。`)
  return value
}
export function chartObjectSheetId(input: Pick<ChartObjectCommandInput, 'sheetId' | 'subUnitId'>): string | undefined {
  const sheetId = identifier(input.sheetId, '工作表标识'), subUnitId = identifier(input.subUnitId, '工作表标识')
  if (sheetId && subUnitId && sheetId !== subUnitId) throw new TypeError('图表命令的工作表标识不一致。')
  return sheetId ?? subUnitId
}
export function assertChartObjectRevision(actual: number, expected?: number): void {
  if (expected === undefined) return
  if (!Number.isSafeInteger(expected) || expected < 0) throw new TypeError('图表命令版本无效。')
  if (actual !== expected) throw new Error('工作簿内容已更新，请重新选择图表。')
}

export function resolveChartObjectTarget(studio: OfficeStudio, input: ChartObjectCommandInput = {}): ChartObjectTarget | undefined {
  if (!input || typeof input !== 'object' || Array.isArray(input)) throw new TypeError('图表命令参数必须为对象。')
  const chartId = identifier(input.chartId, '图表标识'), drawingId = identifier(input.drawingId, '图表标识')
  if (chartId && drawingId && chartId !== drawingId) throw new TypeError('图表与绘图标识不一致。')
  const declaredUnit = identifier(input.unitId, '工作簿标识'), declaredSheet = chartObjectSheetId(input)
  const candidateFocus = chartId || drawingId ? undefined : chartObjectUIHost(studio)?.target()
  const focused = candidateFocus && (!declaredUnit || candidateFocus.unitId === declaredUnit)
    && (!declaredSheet || candidateFocus.sheetId === declaredSheet) ? candidateFocus : undefined
  const unitId = declaredUnit ?? focused?.unitId ?? studio.activeWorkbookId
  const workbook = studio.getWorkbookModel(unitId)
  if (!workbook) throw new Error('图表所属工作簿已关闭。')
  const sheetId = declaredSheet ?? focused?.sheetId ?? workbook.activeSheetId
  const id = chartId ?? drawingId ?? focused?.chartId
    ?? workbook.charts?.find(chart => chart.sheetId === sheetId && isPrimaryChartObject(workbook, chart))?.id
  if (!id) return undefined
  const chart = workbook.charts?.find(chart => chart.id === id && chart.sheetId === sheetId)
  if (!chart) throw new Error('图表已删除或不属于目标工作表。')
  const expected = input.expectedRevision ?? focused?.expectedRevision
  const revision = studio.getWorkbookRevision(unitId)
  assertChartObjectRevision(revision, expected)
  const snapshot = input.snapshot ?? focused?.snapshot
  if (snapshot !== undefined && (typeof snapshot !== 'string' || snapshot !== JSON.stringify(chart))) throw new Error('图表内容已更改，请重新选择。')
  const hostSessionId = identifier(input.hostSessionId, '图表宿主会话') ?? focused?.hostSessionId
  const target = { unitId, sheetId, chartId: id, expectedRevision: revision, snapshot: JSON.stringify(chart),
    ...(hostSessionId ? { hostSessionId } : {}) }
  if (hostSessionId && !chartObjectUIHost(studio, target)) throw new Error('图表宿主已切换，旧菜单操作已停止。')
  return target
}
export function assertChartObjectAction(studio: OfficeStudio, target: ChartObjectTarget, action: ChartObjectAction): ChartModel {
  const current = resolveChartObjectTarget(studio, target)
  if (!current) throw new Error('没有选中的图表。')
  const workbook = studio.getWorkbookModel(current.unitId)!
  const chart = workbook.charts!.find(chart => chart.id === current.chartId)!
  const capability = chartObjectCapabilities(workbook, chart, studio.currentUserId)[action]
  if (capability.disabled || capability.hidden) throw new Error(capability.reason ?? '图表操作不可执行。')
  return chart
}

export function chartObjectShortcut(event: Pick<KeyboardEvent, 'key' | 'ctrlKey' | 'metaKey' | 'altKey' | 'shiftKey' | 'isComposing' | 'keyCode'>): ChartObjectCommandId | typeof PASTE_COPIED_CHART_COMMAND | undefined {
  if (event.isComposing || event.keyCode === 229 || event.altKey) return undefined
  if ((event.ctrlKey || event.metaKey) && !event.shiftKey) {
    if (event.key.toLowerCase() === 'c') return CHART_OBJECT_COMMANDS.copy
    if (event.key.toLowerCase() === 'v') return PASTE_COPIED_CHART_COMMAND
  }
  if (!event.ctrlKey && !event.metaKey && !event.shiftKey && event.key === 'F2') return CHART_OBJECT_COMMANDS.edit
  return undefined
}
