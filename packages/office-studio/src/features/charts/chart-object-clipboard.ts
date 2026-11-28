import { randomId } from '../../api/model-utils'
import type { ChartImageResource, ChartModel, WorkbookModel } from '../../domain/model'
import { cloneModel } from '../../domain/serialization'
import { normalizeDrawingAffineGeometry } from '../../domain/drawing-affine'
import { createChartDefinition } from './chart-engine'
import { resolveChartImageResource } from './chart-image-resources'
import type { ChartObjectPasteInput, ChartObjectTarget } from './chart-object-commands'
import { applyChartSourceImport, assertChartSourceImportSource, planChartSourceImport, type ChartSourceImportPlan } from './chart-source-import'
import { copyPivotReportCache } from '../slicer/slicer-connections-lifecycle'
import { prepareCopiedPivotSourceReport } from '../pivot/services/pivot-source-lifecycle'
import { assertChartReadPermission } from './services/chart-statistics-service'

export const CHART_OBJECT_CLIPBOARD_MAX_BYTES = 32 * 1024 * 1024
export interface ChartObjectClipboard {
  version: 1
  id: string
  source: { unitId: string; sheetId: string; chartId: string; snapshot: string }
  chart: ChartModel
  images: Record<string, ChartImageResource>
}
export function createChartObjectId(): string { return randomId('chart') }
export function captureChartObjectClipboard(workbook: WorkbookModel, target: ChartObjectTarget): ChartObjectClipboard {
  const chart = workbook.charts?.find(chart => chart.id === target.chartId && chart.sheetId === target.sheetId)
  if (!chart || workbook.id !== target.unitId || JSON.stringify(chart) !== target.snapshot) throw new Error('图表复制目标已更新。')
  // The DTO carries the author and its snapshot; reject an oversized author
  // before allocating its detached copy. The final serialized DTO is checked too.
  if (target.snapshot.length * 4 > CHART_OBJECT_CLIPBOARD_MAX_BYTES) throw new RangeError('图表剪贴板超出 32 MiB 预算。')
  const images: Record<string, ChartImageResource> = Object.create(null)
  const maskId = chart.options?.wordCloud?.maskImageId ?? chart.options?.wordCloud?.maskImage
  if (maskId) {
    const image = resolveChartImageResource(workbook, chart, maskId)
    if (!image) throw new Error('图表引用的图片资源已经丢失。')
    images[maskId] = cloneModel(image)
  }
  const result: ChartObjectClipboard = { version: 1, id: randomId('chart-clipboard'),
    source: { unitId: target.unitId, sheetId: target.sheetId, chartId: target.chartId, snapshot: target.snapshot }, chart: cloneModel(chart), images }
  if (JSON.stringify(result).length * 2 > CHART_OBJECT_CLIPBOARD_MAX_BYTES) throw new RangeError('图表剪贴板超出 32 MiB 预算。')
  return result
}
export function chartObjectClipboardContent(clipboard: ChartObjectClipboard): { plain: string; html: string } {
  const token = clipboard.id.replace(/[^A-Za-z0-9:_-]/g, '')
  if (!token || token !== clipboard.id) throw new TypeError('图表剪贴板标识无效。')
  const text = clipboard.chart.title || '图表'
  const escaped = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
  return { plain: `${text}\n[office-chart:${token}]`, html: `<div data-office-chart-clipboard="${token}">${escaped}</div>` }
}
export function chartObjectClipboardIdFromHtml(html: string | undefined): string | undefined {
  return html && /\bdata-office-chart-clipboard\s*=\s*["']([A-Za-z0-9:_-]+)["']/i.exec(html)?.[1] || undefined
}
export function chartObjectClipboardIdFromContent(content: { text?: string; html?: string }): string | undefined {
  const html = chartObjectClipboardIdFromHtml(content.html)
  const plain = content.text && /(?:^|\n)\[office-chart:([A-Za-z0-9:_-]+)\]\s*$/.exec(content.text)?.[1] || undefined
  return html && plain && html !== plain ? undefined : html ?? plain
}

export function assertChartObjectClipboardSource(workbook: WorkbookModel, clipboard: ChartObjectClipboard): void {
  if (!clipboard || clipboard.version !== 1 || !clipboard.id || workbook.id !== clipboard.source.unitId
    || JSON.stringify(clipboard.chart) !== clipboard.source.snapshot) throw new TypeError('图表剪贴板数据或来源工作簿无效。')
  const original = workbook.charts?.find(chart => chart.id === clipboard.source.chartId && chart.sheetId === clipboard.source.sheetId)
  if (!original || JSON.stringify(original) !== clipboard.source.snapshot) throw new Error('复制源图表已更改或删除，请重新复制。')
  for (const [id, image] of Object.entries(clipboard.images)) {
    const current = resolveChartImageResource(workbook, original, id)
    if (!current || JSON.stringify(current) !== JSON.stringify(image)) throw new Error('复制源图片资源已更改，请重新复制图表。')
  }
}
export interface ChartObjectPasteContext {
  sourceWorkbook?: WorkbookModel
  /** UI prepares the same dependency closure asynchronously before committing. */
  sourceImport?: ChartSourceImportPlan
  userId?: string
}

/** Same-unit copies keep live bindings. Cross-unit copies import a detached
 * sparse dependency closure. Failed author/geometry validation never changes
 * the caller's workbook; the returned author is inserted by its transaction. */
export function chartObjectPasteDefinition(workbook: WorkbookModel, clipboard: ChartObjectClipboard,
  input: ChartObjectPasteInput & { unitId: string; sheetId: string; newChartId: string }, context: ChartObjectPasteContext = {}): ChartModel {
  if (!clipboard || clipboard.version !== 1 || !clipboard.id || JSON.stringify(clipboard.chart) !== clipboard.source.snapshot) throw new TypeError('图表剪贴板数据无效。')
  if (workbook.id !== input.unitId) throw new Error('粘贴参数不属于目标工作簿。')
  if (input.clipboardId !== undefined && input.clipboardId !== clipboard.id) throw new Error('图表剪贴板已更新，请重新粘贴。')
  if (typeof input.newChartId !== 'string' || !input.newChartId.trim() || input.newChartId.length > 512) throw new TypeError('粘贴后的图表标识无效。')
  if (workbook.charts?.some(chart => chart.id === input.newChartId)) throw new Error('粘贴后的图表标识已经存在。')
  const crossWorkbook = clipboard.source.unitId !== input.unitId, sourceWorkbook = crossWorkbook ? context.sourceWorkbook : workbook
  if (!sourceWorkbook) throw new Error('跨工作簿粘贴需要可读取的来源工作簿。')
  assertChartObjectClipboardSource(sourceWorkbook, clipboard)
  if (context.userId !== undefined) assertChartReadPermission(sourceWorkbook, clipboard.chart, context.userId)
  if (crossWorkbook && (typeof context.userId !== 'string' || !context.userId.trim())) throw new TypeError('跨工作簿图表粘贴需要明确的用户权限上下文。')
  const candidate: WorkbookModel = { ...workbook, sheets: [...workbook.sheets],
    ...(workbook.pivotCaches ? { pivotCaches: [...workbook.pivotCaches] } : {}) }
  let author = cloneModel(clipboard.chart)
  if (crossWorkbook) {
    const plan = context.sourceImport ?? planChartSourceImport(sourceWorkbook, workbook, clipboard, { userId: context.userId!, targetSheetId: input.sheetId })
    if (plan.sourceUnitId !== sourceWorkbook.id || plan.sourceChartId !== clipboard.source.chartId || plan.sourceSnapshot !== clipboard.source.snapshot
      || plan.targetSheetId !== input.sheetId) throw new TypeError('图表依赖计划与剪贴板来源不匹配。')
    assertChartSourceImportSource(sourceWorkbook, plan, context.userId!)
    applyChartSourceImport(candidate, plan)
    author = cloneModel(plan.chart)
  } else if (author.ownedPivot) {
    const original = clipboard.chart.ownedPivot!
    author.ownedPivot.id = randomId('owned-pivot'); author.ownedPivot.revision = 0
    author.ownedPivot.anchor = { ...author.ownedPivot.anchor, unitId: input.unitId, subUnitId: input.sheetId }
    copyPivotReportCache(workbook, candidate, original, author.ownedPivot, new Map())
    prepareCopiedPivotSourceReport(author.ownedPivot, true)
  }
  const position = input.position ?? { x: clipboard.chart.position.x + 24, y: clipboard.chart.position.y + 24 }
  const source = { ...author.source, range: { ...author.source.range }, sheetId: author.source.sheetId ?? clipboard.chart.sheetId }
  const drawingGeometry = author.drawingGeometry ? normalizeDrawingAffineGeometry({ ...author.drawingGeometry,
    matrix: { ...author.drawingGeometry.matrix, e: author.drawingGeometry.matrix.e + position.x - clipboard.chart.position.x,
      f: author.drawingGeometry.matrix.f + position.y - clipboard.chart.position.y } }) : undefined
  const created = createChartDefinition(candidate, input.sheetId, { ...author, source, position: { ...position }, drawingGeometry }, input.newChartId)
  if (context.userId !== undefined) assertChartReadPermission(candidate, created, context.userId)
  if (crossWorkbook) {
    workbook.sheets = candidate.sheets; workbook.tables = candidate.tables; workbook.pivotTables = candidate.pivotTables
    workbook.styles = candidate.styles; workbook.rangeThemes = candidate.rangeThemes; workbook.metadata = candidate.metadata
  }
  if (crossWorkbook || author.ownedPivot) workbook.pivotCaches = candidate.pivotCaches
  return created
}
