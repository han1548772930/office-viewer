import type { ChartModel } from '../../domain/model'
import type { DrawCommand, RenderRect } from '../../rendering/contracts'
import { measureSpreadsheetText as measureCanvasText } from '../../rendering/spreadsheet/text-metrics'
import { sheetTextGraphemes } from '../sheet-workbench/sheet-cell-text-layout'
import type { ChartData, ChartLayout } from './chart-engine'
import { chartHierarchyPointStyle } from './chart-derived'
import { chartColorOpacity, chartFont, chartRect } from './chart-geometry'
import { chartFillPatternPaint } from './chart-fill-pattern'
import { chartTreemapBackgroundPaint, chartTreemapLabelBackground, chartTreemapLabelText, chartTreemapParentLayout,
  chartTreemapTextBlock, chartTreemapTextColor, resolveChartTreemapLabel } from './chart-treemap'

export function drawChartTreemap(commands: DrawCommand[], data: ChartData, chart: ChartModel, layout: ChartLayout,
  color: string, palette: string[], background = '#ffffff'): void {
  const rectangles = layout.treemap
  if (!rectangles) throw new TypeError('树图绘制需要匹配的完整层级布局。')
  const source = data.series[0]!, mode = chartTreemapParentLayout(chart.options?.treemap)
  for (const rect of rectangles) {
    const style = chartHierarchyPointStyle(chart, source.field, rect, palette, source.sourceName ?? source.name)
    chartRect(commands, rect, chartColorOpacity(style.color!, style.fillOpacity ?? 1),
      { ...style.border, color: style.border?.color ?? chart.options?.appearance?.background ?? background, width: style.border?.width ?? (rect.depth === 0 ? 2 : 1) },
      layout.plot, chartFillPatternPaint(style.fillPattern, style.fillOpacity ?? 1))
  }
  // Parent text in overlapping mode must paint above every descendant label.
  const labelRectangles = [...rectangles].sort((left, right) => right.depth - left.depth)
  for (const rect of labelRectangles) {
    const parent = Boolean(rect.renderNode.children?.length), style = chartHierarchyPointStyle(chart, source.field, rect, palette, source.sourceName ?? source.name)
    const label = resolveChartTreemapLabel(chart, style.label)
    if (parent ? mode === 'none' : !label.visible) continue
    const text = parent ? rect.node.name : chartTreemapLabelText(rect.renderNode, label, data.date1904)
    const font = chartFont(label, 12), lineHeight = parent ? label.fontSize : label.fontSize * 1.2
    const available = { width: Math.max(0, rect.width - 4), height: parent ? Math.min(rect.height, rect.labelHeight) : Math.max(0, rect.height - 4) }
    const block = chartTreemapTextBlock(text, available, font, lineHeight, measureCanvasText, value => sheetTextGraphemes(value).map(glyph => glyph.text))
    if (!block) continue
    const centerX = rect.x + rect.width / 2, centerY = parent ? rect.y + Math.min(rect.height, rect.labelHeight) / 2 : rect.y + rect.height / 2
    const box: RenderRect = { x: centerX - block.width / 2, y: centerY - block.height / 2, width: block.width, height: block.height }
    const fill = parent ? undefined : chartTreemapLabelBackground(label.backgroundColor)
    if (fill !== undefined) chartRect(commands, box, chartColorOpacity(fill), undefined, rect, undefined, chartTreemapBackgroundPaint(fill, box))
    const textColor = chartTreemapTextColor(style.color!, label.color, color)
    for (let line = 0; line < block.lines.length; line++) commands.push({ kind: 'text', x: centerX, y: box.y + (line + 0.5) * block.lineHeight,
      text: block.lines[line]!, font, color: textColor, align: 'center', baseline: 'middle', clip: rect })
  }
}
