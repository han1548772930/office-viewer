import { randomId } from '../../api/model-utils'
import type { ChartImageResource, ChartModel, WorkbookModel } from '../../domain/model'
import { cloneModel } from '../../domain/serialization'
import { createChartDefinition } from './chart-engine'
import { readChartImageResources, resolveChartImageResource } from './chart-image-resources'
import type { ChartObjectPasteInput, ChartObjectTarget } from './chart-object-commands'

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

/** Copy preserves live source bindings in one workbook. The clipboard never
 * stores a second author model, derived arrays, hover state or image bitmap. */
export function chartObjectPasteDefinition(workbook: WorkbookModel, clipboard: ChartObjectClipboard,
  input: ChartObjectPasteInput & { unitId: string; sheetId: string; newChartId: string }): ChartModel {
  if (!clipboard || clipboard.version !== 1 || !clipboard.id || JSON.stringify(clipboard.chart) !== clipboard.source.snapshot) throw new TypeError('图表剪贴板数据无效。')
  if (workbook.id !== input.unitId || clipboard.source.unitId !== input.unitId) throw new Error('跨工作簿图表粘贴还需导入源工作表与资源依赖；当前剪贴板内容已保留。')
  if (input.clipboardId !== undefined && input.clipboardId !== clipboard.id) throw new Error('图表剪贴板已更新，请重新粘贴。')
  if (typeof input.newChartId !== 'string' || !input.newChartId.trim() || input.newChartId.length > 512) throw new TypeError('粘贴后的图表标识无效。')
  if (workbook.charts?.some(chart => chart.id === input.newChartId)) throw new Error('粘贴后的图表标识已经存在。')
  const original = workbook.charts?.find(chart => chart.id === clipboard.source.chartId && chart.sheetId === clipboard.source.sheetId)
  if (!original || JSON.stringify(original) !== clipboard.source.snapshot) throw new Error('复制源图表已更改或删除，请重新复制。')
  const resources = readChartImageResources(workbook)
  for (const [id, image] of Object.entries(clipboard.images)) if (!resources[id] || JSON.stringify(resources[id]) !== JSON.stringify(image)) throw new Error('复制源图片资源已更改，请重新复制图表。')
  const position = input.position ?? { x: clipboard.chart.position.x + 24, y: clipboard.chart.position.y + 24 }
  const source = { ...clipboard.chart.source, range: { ...clipboard.chart.source.range },
    sheetId: clipboard.chart.source.sheetId ?? clipboard.chart.sheetId }
  const created = createChartDefinition(workbook, input.sheetId, { ...cloneModel(clipboard.chart), source, position: { ...position } }, input.newChartId)
  // Chart image resources use immutable IDs in this workbook. The new author
  // reference participates in the existing pruning and snapshot/history path.
  return created
}
