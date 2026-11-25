import type { DrawCommand } from '../../rendering/contracts'
import type { ChartModel } from '../../domain/model'
import { measureSpreadsheetText as measureCanvasText } from '../../rendering/spreadsheet/text-metrics'
import type { ChartData, ChartLayout } from './chart-engine'
import { chartHierarchyPointStyle } from './chart-derived'
import { chartColorOpacity, chartFont, chartPolygon, type ChartVector } from './chart-geometry'
import { chartFillPatternPaint } from './chart-fill-pattern'
import { chartLabelBackground, chartLabelBackgroundPaint } from './chart-label-background'
import { drawChartSector } from './chart-sector-draw'
import { chartSunburstCenterText, chartSunburstLabelAnchor, chartSunburstLabelFits, chartSunburstLabelText, chartSunburstLabelVisible, chartSunburstSectorRect,
  chartSunburstTotal, resolveChartSunburstCenterLabel, resolveChartSunburstLabel } from './chart-sunburst'

/** Both native and common Canvas consume the same semantic hierarchy and
 * prepared ring geometry. Labels keep full content; they are never compressed. */
export function drawChartSunburst(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout,
  color: string, palette: string[], background = '#ffffff'): void {
  const sectors = layout.sunburst
  if (!sectors) throw new TypeError('旭日图绘制需要匹配的完整层级布局。')
  const center = { x: layout.plot.x + layout.plot.width / 2, y: layout.plot.y + layout.plot.height / 2 }, total = chartSunburstTotal(data.hierarchy ?? [])
  const source = data.series[0]!
  for (const sector of sectors) {
    const style = chartHierarchyPointStyle(chart, source.field, sector, palette, source.sourceName ?? source.name)
    // A signed parent's child can exceed a full turn. The locked Sector shape
    // paints one circle; cap only physical tessellation, retaining raw angles.
    drawChartSector(commands, center, sector.inner, sector.outer, sector.start, Math.min(sector.end, sector.start + Math.PI * 2),
      chartColorOpacity(style.color!, style.fillOpacity ?? 1), { color: background, width: 2, ...style.border }, layout.plot,
      chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1))
  }
  for (const sector of sectors) {
    const style = chartHierarchyPointStyle(chart, source.field, sector, palette, source.sourceName ?? source.name), label = resolveChartSunburstLabel(chart, style.label)
    if (!chartSunburstLabelVisible(sector.renderNode, total, label)) continue
    const text = chartSunburstLabelText(sector.renderNode, label, data.date1904)
    if (!text) continue
    const font = chartFont(label, 12), lines = text.split('\n'), lineHeight = label.fontSize * 1.2
    const width = lines.reduce((maximum, line) => Math.max(maximum, measureCanvasText(line, font)), 0), height = lines.length * lineHeight
    const anchor = chartSunburstLabelAnchor(center, sector, label.position), radians = anchor.rotation * Math.PI / 180
    const cosine = Math.cos(radians), sine = Math.sin(radians), rect = chartSunburstSectorRect(center, sector)
    if (!chartSunburstLabelFits(rect, { width: Math.abs(cosine) * width + Math.abs(sine) * height, height: Math.abs(sine) * width + Math.abs(cosine) * height })) continue
    const left = anchor.align === 'left' ? 0 : anchor.align === 'right' ? -width : -width / 2, top = -height / 2
    const transform = (point: ChartVector): ChartVector => ({ x: anchor.x + point.x * cosine - point.y * sine, y: anchor.y + point.x * sine + point.y * cosine })
    const fill = chartLabelBackground(label.backgroundColor)
    if (fill !== undefined && width > 0) {
      const paint = chartLabelBackgroundPaint(fill, { x: left, y: top, width, height })
      const gradient = paint && { ...paint, bounds: { x: 0, y: 0, width: 1, height: 1 },
        start: transform({ x: left + paint.start.x * width, y: top + paint.start.y * height }),
        end: transform({ x: left + paint.end.x * width, y: top + paint.end.y * height }) }
      chartPolygon(commands, [transform({ x: left, y: top }), transform({ x: left + width, y: top }),
        transform({ x: left + width, y: top + height }), transform({ x: left, y: top + height })], chartColorOpacity(fill), undefined, layout.frame, undefined, gradient)
    }
    for (let line = 0; line < lines.length; line++) commands.push({ kind: 'text', ...transform({ x: 0, y: top + (line + 0.5) * lineHeight }),
      text: lines[line]!, font, color: label.color ?? color, align: anchor.align, baseline: 'middle', rotation: anchor.rotation, clip: layout.frame })
  }
  const centerLabel = resolveChartSunburstCenterLabel(chart)
  if (centerLabel.visible) commands.push({ kind: 'text', ...center, text: chartSunburstCenterText(data.hierarchy ?? [], centerLabel, data.date1904),
    font: chartFont(centerLabel, 16), color: centerLabel.color ?? color, align: 'center', baseline: 'middle', clip: layout.frame })
}
