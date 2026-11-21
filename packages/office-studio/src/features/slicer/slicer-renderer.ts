import type { SlicerModel, TimelineLevel } from '../../domain/model'
import type { DrawCommand, RenderRect } from '../../rendering/contracts'
import { SpreadsheetGeometry, type GridViewport } from '../../rendering/spreadsheet/layout'
import { parseTimelineDate, timelineDateToInput, type TimelineDate } from './timeline-date'
import { slicerCanvasIcon, type SlicerIconName } from './slicer-canvas-icons'

export interface SlicerCanvasControl { model: SlicerModel; items: string[] }
export interface SlicerCanvasPeriod { key: number; start: number; end: number; label: string; groupLabel: string }
export interface SlicerObjectLayout {
  id: string
  model: SlicerModel
  rect: RenderRect
  header: RenderRect
  clear: RenderRect
  multiSelect: RenderRect
  items: Array<{ value: string; rect: RenderRect; selected: boolean; index: number }>
  periods: Array<{ period: SlicerCanvasPeriod; rect: RenderRect; selected: boolean }>
  track?: RenderRect
  level: RenderRect
  scrollbar?: RenderRect
  scrollThumb?: RenderRect
  selectionLabel?: string
  selectionLabelRect?: RenderRect
  content?: RenderRect
  groupLabelBand?: RenderRect
  periodLabelBand?: RenderRect
  startHandle?: RenderRect
  endHandle?: RenderRect
  scrollOffset: number
  maxScrollOffset: number
  visibleCount: number
  menu?: RenderRect[]
}

export type SlicerCanvasHit =
  | { type: 'body' | 'move' | 'resize' | 'clear' | 'multiSelect' | 'timelineLevel' | 'timelineStart' | 'timelineEnd' | 'timelineMenu'; id: string }
  | { type: 'scroll'; id: string; part: 'track' | 'thumb' }
  | { type: 'item'; id: string; value: string; index: number }
  | { type: 'period'; id: string; period: SlicerCanvasPeriod; index: number }

const TIMELINE_SLOT = 64
const TIMELINE_LEVELS: TimelineLevel[] = ['year', 'quarter', 'month', 'day']

export function slicerObjectLayouts(geometry: SpreadsheetGeometry, controls: readonly SlicerCanvasControl[], date1904: boolean, locale = 'zh-CN'): SlicerObjectLayout[] {
  return controls.map(({ model, items }) => {
    const x = model.layout?.left ?? 44 + geometry.columns.offset(model.position.column)
    const y = model.layout?.top ?? 26 + geometry.rows.offset(model.position.row)
    const width = model.layout?.width ?? geometry.columns.offset(model.position.column + model.size.columns) - geometry.columns.offset(model.position.column)
    const height = model.layout?.height ?? geometry.rows.offset(model.position.row + model.size.rows) - geometry.rows.offset(model.position.row)
    const rect = { x, y, width, height }
    const headerHeight = model.type === 'timeline'
      ? model.showHeader === false ? 0 : Math.min(32, height)
      : model.showCaption === false ? 0 : Math.min(32, height)
    const header = { x, y, width, height: headerHeight }
    const clear = { x: x + width - (model.type === 'timeline' ? 25 : 24), y: y + (model.type === 'timeline' ? 5 : 6), width: 20, height: 20 }
    const multiSelect = { x: x + width - 46, y: y + 6, width: 20, height: 20 }
    const level = model.type === 'timeline'
      ? { x: x + width - Math.min(100, width / 3), y: y + headerHeight, width: Math.min(100, width / 3), height: Math.min(24, Math.max(0, height - headerHeight)) }
      : { x: x + width - 68, y: y + 6, width: 62, height: 22 }
    if (model.type === 'timeline') return timelineLayout(model, items, rect, header, clear, level, date1904, locale)

    const rowHeight = model.rowHeight ?? 24
    const columnCount = Math.max(1, model.columnCount ?? 1)
    const contentY = headerHeight
    const innerHeight = Math.max(0, height - contentY - 8)
    let visibleRows = Math.max(1, Math.ceil(innerHeight / (rowHeight + 4)))
    while (visibleRows > 1 && (innerHeight - rowHeight) / (visibleRows - 1) - rowHeight < 0) visibleRows -= 1
    const pitch = visibleRows > 1 ? (innerHeight - rowHeight) / (visibleRows - 1) : rowHeight
    const totalRows = Math.ceil(items.length / columnCount)
    const maxScrollRows = Math.max(0, totalRows - visibleRows)
    const scrollRow = Math.min(maxScrollRows, Math.floor((model.startItem ?? 0) / columnCount))
    const firstItem = scrollRow * columnCount
    const scrollWidth = totalRows > visibleRows ? 14 : 0
    const availableWidth = width - 8 - scrollWidth
    const itemWidth = Math.max(0, (availableWidth - 4 * (columnCount - 1)) / columnCount)
    const visibleItems = items.slice(firstItem, Math.min(items.length, firstItem + visibleRows * columnCount))
    const layoutItems = visibleItems.map((value, visibleIndex) => {
      const row = Math.floor(visibleIndex / columnCount)
      const column = visibleIndex % columnCount
      const filterItems = model.selectedItems
      return {
        value,
        rect: { x: x + 4 + column * (itemWidth + 4), y: y + contentY + 4 + row * pitch, width: itemWidth, height: rowHeight },
        selected: filterItems == null || filterItems.includes(value),
        index: firstItem + visibleIndex,
      }
    })
    const scrollbar = scrollWidth ? { x: x + width - 14, y: y + contentY + 4, width: 10, height: Math.max(0, height - contentY - 8) } : undefined
    const thumbHeight = scrollbar ? Math.min(scrollbar.height, Math.max(12, scrollbar.height * visibleRows / totalRows)) : 0
    const scrollThumb = scrollbar ? { ...scrollbar, y: scrollbar.y + (maxScrollRows ? (scrollbar.height - thumbHeight) * scrollRow / maxScrollRows : 0), height: thumbHeight } : undefined
    return {
      id: model.id, model, rect, header, clear, multiSelect, level,
      items: layoutItems, periods: [], scrollbar, scrollThumb,
      content: { x, y: y + contentY, width, height: Math.max(0, height - contentY) },
      scrollOffset: firstItem, maxScrollOffset: maxScrollRows * columnCount, visibleCount: visibleRows * columnCount,
    }
  })
}

function timelineLayout(model: SlicerModel, items: readonly string[], rect: RenderRect, header: RenderRect, clear: RenderRect, level: RenderRect, date1904: boolean, locale: string): SlicerObjectLayout {
  const { x, y, width, height } = rect
  const secondaryHeight = model.showSelectionLabel === false && model.showTimeLevel === false ? 0 : Math.min(24, height - header.height)
  const scrollHeight = model.showHorizontalScrollbar === false ? 0 : Math.min(12, height - header.height - secondaryHeight)
  const axisHeight = Math.max(0, height - header.height - secondaryHeight - scrollHeight)
  const axisY = y + header.height + secondaryHeight
  const trackHeight = Math.min(26, axisHeight)
  const labelHeight = Math.min(22, Math.max(0, axisHeight - trackHeight))
  const groupHeight = Math.min(18, Math.max(0, axisHeight - trackHeight - labelHeight))
  const bandY = axisY + Math.max(0, (axisHeight - trackHeight - labelHeight - groupHeight) / 2)
  const groupLabelBand = { x, y: bandY, width, height: groupHeight }
  const periodLabelBand = { x, y: bandY + groupHeight, width, height: labelHeight }
  const axisTrack = { x, y: bandY + groupHeight + labelHeight, width, height: trackHeight }
  const visibleCount = Math.max(1, Math.ceil(width / TIMELINE_SLOT))
  const bounds = timelineBounds(items, date1904)
  const currentLevel = model.level ?? 'month'
  const total = bounds ? levelIndex(bounds.last, currentLevel) - levelIndex(bounds.first, currentLevel) + 1 : 0
  const maxScrollOffset = Math.max(0, total - visibleCount)
  const first = clamp(bounds ? timelineOffsetAtSerial(bounds, model.scrollPosition ?? bounds.first.serial, currentLevel, date1904) : 0, 0, maxScrollOffset)
  const firstVisible = Math.max(0, Math.floor(first) - 1)
  const shown = bounds ? periodWindow(bounds, currentLevel, date1904, firstVisible, visibleCount + 2, locale) : []
  const selection = model.timelineRange
  const selectedStart = selection?.start == null ? undefined : parseTimelineDate(selection.start, date1904)?.serial
  const selectedEnd = selection?.end == null ? undefined : parseTimelineDate(selection.end, date1904)?.serial
  const periodLayouts = shown.map((period, index) => {
    const absoluteIndex = firstVisible + index
    const slotX = x + absoluteIndex * TIMELINE_SLOT - first * TIMELINE_SLOT
    return {
      period,
      rect: { x: slotX, y: bandY, width: TIMELINE_SLOT, height: groupHeight + labelHeight + trackHeight },
      selected: selection != null && (selectedStart === undefined || period.end >= selectedStart) && (selectedEnd === undefined || period.start <= selectedEnd),
    }
  })
  const scrollTrack = scrollHeight ? { x: x + 5, y: y + height - scrollHeight, width: Math.max(0, width - 10), height: scrollHeight } : undefined
  const thumbWidth = scrollTrack ? Math.min(scrollTrack.width, Math.max(16, scrollTrack.width * visibleCount / Math.max(visibleCount, total))) : 0
  const scrollThumb = scrollTrack ? { ...scrollTrack, x: scrollTrack.x + (total > visibleCount ? (scrollTrack.width - thumbWidth) * first / (total - visibleCount) : 0), width: thumbWidth } : undefined
  const menuRowHeight = Math.min(24, Math.max(0, y + height - level.y - level.height) / 4)
  const menu = TIMELINE_LEVELS.map((_, index) => ({ x: level.x, y: level.y + level.height + index * menuRowHeight, width: level.width, height: menuRowHeight }))
  const rangeLabel = selection ? `${timelineDateToInput(selection.start, date1904)} – ${timelineDateToInput(selection.end, date1904)}` : '所有时段'
  const selected = periodLayouts.filter(period => period.selected)
  const handleWidth = Math.min(8, width)
  const startHandle = selected.length ? { x: clamp(selected[0]!.rect.x - handleWidth / 2, x, x + width - handleWidth), y: axisTrack.y, width: handleWidth, height: axisTrack.height } : undefined
  const endHandle = selected.length ? { x: clamp(selected[selected.length - 1]!.rect.x + TIMELINE_SLOT - handleWidth / 2, x, x + width - handleWidth), y: axisTrack.y, width: handleWidth, height: axisTrack.height } : undefined
  return { id: model.id, model, rect, header, clear, multiSelect: { x, y, width: 0, height: 0 }, items: [], periods: periodLayouts, track: axisTrack, groupLabelBand, periodLabelBand, level, scrollbar: scrollTrack, scrollThumb, selectionLabel: model.showSelectionLabel === false ? undefined : rangeLabel, selectionLabelRect: { x, y: y + header.height, width: Math.max(0, width - (model.showTimeLevel === false ? 0 : level.width)), height: secondaryHeight }, menu, startHandle, endHandle, scrollOffset: first, maxScrollOffset, visibleCount }
}

export interface SlicerDrawState {
  timelineMenuId?: string
  selectedSlicerId?: string
  multiSelect?: ReadonlySet<string>
  focusedItems?: ReadonlyMap<string, number>
  focusedPeriods?: ReadonlyMap<string, number>
  onIconReady?: () => void
}

export function slicerObjectDrawCommands(layouts: readonly SlicerObjectLayout[], state: SlicerDrawState = {}): DrawCommand[] {
  const commands: DrawCommand[] = []
  for (const layout of layouts) {
    const colors = slicerPalette(layout.model.style, layout.model.type === 'timeline')
    const begin = commands.length
    commands.push({ kind: 'rect', ...layout.rect, fill: colors.background })
    if (layout.header.height > 0) {
      commands.push({ kind: 'rect', ...layout.header, fill: colors.header, stroke: colors.headerBorder, strokeWidth: 1 })
      commands.push({ kind: 'text', x: layout.rect.x + 8, y: layout.header.y + layout.header.height / 2, text: layout.model.caption, font: '600 14px Arial, sans-serif', color: colors.headerText, baseline: 'middle', maxWidth: Math.max(0, layout.rect.width - 60) })
      const filtered = layout.model.type === 'timeline' ? Boolean(layout.model.timelineRange) : layout.model.selectedItems !== undefined
      drawIcon(commands, layout.clear, 'clear', colors.headerText, state.onIconReady, filtered ? 1 : .48)
      if (layout.model.type === 'slicer') {
        const active = state.multiSelect?.has(layout.id)
        if (active) roundedRect(commands, layout.multiSelect, 6, mix(colors.background, colors.selected, .16))
        drawIcon(commands, layout.multiSelect, 'multi', colors.headerText, state.onIconReady, active ? 1 : .48)
      }
    }
    if (layout.model.type === 'slicer') {
      for (const item of layout.items) {
        const focused = state.focusedItems?.get(layout.id) === item.index
        const lineWidth = focused ? 1.25 : .75
        roundedRect(commands, item.rect, 8, colors.itemBorder)
        roundedRect(commands, { x: item.rect.x + lineWidth, y: item.rect.y + lineWidth, width: Math.max(0, item.rect.width - lineWidth * 2), height: Math.max(0, item.rect.height - lineWidth * 2) }, Math.max(0, 8 - lineWidth), item.selected ? colors.selected : colors.item)
        commands.push({ kind: 'text', x: item.rect.x + 8, y: item.rect.y + item.rect.height / 2, text: item.value, font: '14px Arial, sans-serif', color: item.selected ? colors.selectedText : colors.text, baseline: 'middle', maxWidth: Math.max(0, item.rect.width - 16), clip: layout.content })
      }
    } else {
      if (layout.selectionLabel && layout.selectionLabelRect) {
        const label = layout.selectionLabelRect
        commands.push({ kind: 'text', x: label.x + 5, y: label.y + label.height / 2, text: layout.selectionLabel, font: '600 12px Arial, sans-serif', color: colors.selectionText, baseline: 'middle', maxWidth: Math.max(0, label.width - 10) })
      }
      if (layout.model.showTimeLevel !== false) {
        commands.push({ kind: 'text', x: layout.level.x + 5, y: layout.level.y + layout.level.height / 2, text: levelLabel(layout.model.level ?? 'month'), font: '12px Arial, sans-serif', color: colors.text, baseline: 'middle', maxWidth: Math.max(0, layout.level.width - 28) })
        drawIcon(commands, { x: layout.level.x + layout.level.width - 17, y: layout.level.y + (layout.level.height - 16) / 2, width: 16, height: 16 }, state.timelineMenuId === layout.id ? 'moreUp' : 'moreDown', colors.text, state.onIconReady)
      }
      const group = layout.groupLabelBand
      const label = layout.periodLabelBand
      const track = layout.track
      if (group && label && track) {
        const clip = { x: layout.rect.x, y: group.y, width: layout.rect.width, height: group.height + label.height + track.height }
        for (let index = 0; index < layout.periods.length; index += 1) {
          const period = layout.periods[index]!
          const px = period.rect.x
          commands.push({ kind: 'rect', x: px, y: label.y, width: period.rect.width, height: label.height, fill: index % 2 ? colors.periodAlternate : colors.background, clip })
          commands.push({ kind: 'text', x: px + period.rect.width / 2, y: label.y + label.height / 2, text: period.period.label, font: '12px Arial, sans-serif', color: colors.periodText, align: 'center', baseline: 'middle', maxWidth: Math.max(0, period.rect.width - 10), clip })
          commands.push({ kind: 'rect', x: px, y: track.y, width: period.rect.width, height: track.height, fill: period.selected ? colors.selected : colors.item, clip })
          commands.push({ kind: 'rect', x: px + period.rect.width - 1, y: track.y, width: 1, height: track.height, fill: colors.background, clip })
          if (period.period.groupLabel && (!index || layout.periods[index - 1]!.period.groupLabel !== period.period.groupLabel)) commands.push({ kind: 'text', x: px + 5, y: group.y + group.height / 2, text: period.period.groupLabel, font: '10px Arial, sans-serif', color: colors.periodText, baseline: 'middle', clip })
          if (state.focusedPeriods?.get(layout.id) === period.period.key) commands.push({ kind: 'rect', x: px + 1, y: label.y + 1, width: Math.max(0, period.rect.width - 2), height: Math.max(0, label.height + track.height - 2), fill: 'transparent', stroke: colors.border, strokeWidth: 1, clip })
        }
        commands.push({ kind: 'rect', x: track.x, y: track.y, width: track.width, height: 1, fill: mix(colors.background, colors.border, .36) })
        if (layout.startHandle && layout.endHandle) {
          const first = layout.startHandle.x + layout.startHandle.width / 2
          const last = layout.endHandle.x + layout.endHandle.width / 2
          commands.push({ kind: 'line', x1: first, y1: track.y, x2: last, y2: track.y, color: colors.itemBorder, width: 1 })
          for (const handle of [layout.startHandle, layout.endHandle]) commands.push({ kind: 'rect', x: handle.x + handle.width / 2, y: handle.y, width: 1, height: handle.height, fill: colors.itemBorder })
        }
      }
    }
    if (layout.scrollbar && layout.scrollThumb) {
      commands.push({ kind: 'rect', ...layout.scrollbar, fill: colors.background, stroke: colors.border, strokeWidth: 1 })
      if (layout.model.type === 'timeline') {
        roundedRect(commands, { x: layout.scrollThumb.x + 1, y: layout.scrollThumb.y + Math.max(1, (layout.scrollThumb.height - 4) / 2), width: Math.max(0, layout.scrollThumb.width - 2), height: Math.min(4, Math.max(0, layout.scrollThumb.height - 2)) }, 2, colors.border)
      } else commands.push({ kind: 'rect', x: layout.scrollThumb.x + 1, y: layout.scrollThumb.y + 1, width: Math.max(0, layout.scrollThumb.width - 2), height: Math.max(0, layout.scrollThumb.height - 2), fill: colors.border })
    }
    if (state.timelineMenuId === layout.id) layout.menu?.forEach((menuItem, index) => {
      const selected = TIMELINE_LEVELS[index] === (layout.model.level ?? 'month')
      commands.push({ kind: 'rect', ...menuItem, fill: selected ? mix(colors.background, colors.selected, .16) : colors.background, stroke: colors.border, strokeWidth: .5 })
      if (selected) commands.push({ kind: 'rect', x: menuItem.x, y: menuItem.y + 3, width: 3, height: Math.max(0, menuItem.height - 6), fill: colors.selected })
      commands.push({ kind: 'text', x: menuItem.x + 5, y: menuItem.y + menuItem.height / 2, text: levelLabel(TIMELINE_LEVELS[index]!), font: '12px Arial, sans-serif', color: colors.text, baseline: 'middle' })
    })
    if (state.selectedSlicerId === layout.id && !layout.model.lockedPosition) {
      const right = layout.rect.x + layout.rect.width - 3
      const bottom = layout.rect.y + layout.rect.height - 3
      for (let index = 0; index < 3; index += 1) commands.push({ kind: 'line', x1: right - 3 - index * 3, y1: bottom, x2: right, y2: bottom - 3 - index * 3, color: colors.border, width: 1 })
    }
    commands.push({ kind: 'rect', x: layout.rect.x + .5, y: layout.rect.y + .5, width: Math.max(0, layout.rect.width - 1), height: Math.max(0, layout.rect.height - 1), fill: 'transparent', stroke: colors.border, strokeWidth: 1 })
    for (let index = begin; index < commands.length; index += 1) commands[index]!.clip ??= layout.rect
  }
  return commands
}

function drawIcon(commands: DrawCommand[], rect: RenderRect, kind: SlicerIconName, color: string, onReady?: () => void, opacity = 1): void {
  const source = slicerCanvasIcon(kind, color, onReady)
  if (source) commands.push({ kind: 'image', ...rect, source, opacity })
}

function roundedRect(commands: DrawCommand[], rect: RenderRect, radius: number, fill: string): void {
  if (rect.width <= 0 || rect.height <= 0) return
  const r = Math.min(radius, rect.width / 2, rect.height / 2)
  const points: Array<{ x: number; y: number }> = []
  const corners = [[rect.x + rect.width - r, rect.y + r], [rect.x + rect.width - r, rect.y + rect.height - r], [rect.x + r, rect.y + rect.height - r], [rect.x + r, rect.y + r]]
  for (let corner = 0; corner < corners.length; corner += 1) {
    for (let step = 0; step <= 6; step += 1) {
      const angle = (corner - 1) * Math.PI / 2 + step * Math.PI / 12
      points.push({ x: corners[corner]![0]! + Math.cos(angle) * r, y: corners[corner]![1]! + Math.sin(angle) * r })
    }
  }
  commands.push({ kind: 'polygon', points, fill })
}

export function hitTestSlicerObjects(layouts: readonly SlicerObjectLayout[], x: number, y: number, timelineMenuId = ''): SlicerCanvasHit | undefined {
  for (const layout of [...layouts].reverse()) {
    if (timelineMenuId === layout.id && layout.menu?.some((rect) => contains(rect, x, y))) return { type: 'timelineMenu', id: layout.id }
    if (!contains(layout.rect, x, y)) continue
    if (layout.header.height && contains(layout.clear, x, y)) return { type: 'clear', id: layout.id }
    if (layout.model.type === 'slicer' && layout.header.height && contains(layout.multiSelect, x, y)) return { type: 'multiSelect', id: layout.id }
    if (layout.model.type === 'timeline') {
      if (layout.model.showTimeLevel !== false && contains(layout.level, x, y)) return { type: 'timelineLevel', id: layout.id }
      if (layout.startHandle && contains(layout.startHandle, x, y)) return { type: 'timelineStart', id: layout.id }
      if (layout.endHandle && contains(layout.endHandle, x, y)) return { type: 'timelineEnd', id: layout.id }
    }
    const itemIndex = layout.items.findIndex((item) => contains(item.rect, x, y))
    if (itemIndex >= 0) return { type: 'item', id: layout.id, value: layout.items[itemIndex]!.value, index: layout.items[itemIndex]!.index }
    const periodIndex = layout.periods.findIndex((period) => contains(period.rect, x, y))
    if (periodIndex >= 0) return { type: 'period', id: layout.id, period: layout.periods[periodIndex]!.period, index: layout.periods[periodIndex]!.period.key }
    if (layout.scrollbar && contains(layout.scrollbar, x, y)) return { type: 'scroll', id: layout.id, part: layout.scrollThumb && contains(layout.scrollThumb, x, y) ? 'thumb' : 'track' }
    if (!layout.model.lockedPosition && contains({ x: layout.rect.x + layout.rect.width - 10, y: layout.rect.y + layout.rect.height - 10, width: 10, height: 10 }, x, y)) return { type: 'resize', id: layout.id }
    if (layout.header.height > 0 && y <= layout.header.y + layout.header.height) return { type: 'move', id: layout.id }
    return { type: 'body', id: layout.id }
  }
  return undefined
}

export function slicerObjectContentPoint(x: number, y: number, viewport: GridViewport): { x: number; y: number } {
  return { x: (x + viewport.scrollLeft) / viewport.zoom, y: (y + viewport.scrollTop) / viewport.zoom }
}

export function timelinePeriods(items: readonly string[], level: TimelineLevel, date1904: boolean, startIndex = 0, count = Number.POSITIVE_INFINITY, locale = 'zh-CN'): SlicerCanvasPeriod[] {
  const bounds = timelineBounds(items, date1904)
  return bounds ? periodWindow(bounds, level, date1904, startIndex, count, locale) : []
}

interface TimelineBounds { first: TimelineDate; last: TimelineDate }
function timelineBounds(items: readonly string[], date1904: boolean): TimelineBounds | undefined {
  let first: TimelineDate | undefined
  let last: TimelineDate | undefined
  for (const value of items) {
    const numericValue = Number(value)
    const date = value.trim() && Number.isFinite(numericValue) ? parseTimelineDate(numericValue, date1904) : parseTimelineDate(value, date1904)
    if (!date) continue
    if (!first || date.serial < first.serial) first = date
    if (!last || date.serial > last.serial) last = date
  }
  return first && last ? { first, last } : undefined
}

function periodWindow(bounds: TimelineBounds, level: TimelineLevel, date1904: boolean, startIndex: number, count: number, locale: string): SlicerCanvasPeriod[] {
  const firstIndex = levelIndex(bounds.first, level)
  const total = levelIndex(bounds.last, level) - firstIndex + 1
  const periods: SlicerCanvasPeriod[] = []
  const from = Math.max(0, Math.floor(startIndex))
  const to = Math.min(total, from + Math.max(0, Math.ceil(count)))
  for (let index = from; index < to; index += 1) {
    const key = firstIndex + index
    const serial = periodSerial(key, level, date1904)
    const date = parseTimelineDate(serial, date1904)
    if (!date) continue
    const label = level === 'year' ? String(date.year)
      : level === 'quarter' ? `Q${Math.floor((date.month - 1) / 3) + 1}`
        : level === 'month' ? new Intl.DateTimeFormat(locale, { month: 'short', timeZone: 'UTC' }).format(calendarDate(date))
          : String(date.day)
    const nextSerial = periodSerial(key + 1, level, date1904)
    const groupLabel = level === 'day' ? new Intl.DateTimeFormat(locale, { month: 'short', year: 'numeric', timeZone: 'UTC' }).format(calendarDate(date)) : level === 'year' ? '' : String(date.year)
    periods.push({ key, start: serial, end: nextSerial - 1, label, groupLabel })
  }
  return periods
}

function periodSerial(index: number, level: TimelineLevel, date1904: boolean): number {
  if (level === 'day') return index
  const year = level === 'year' ? index : Math.floor(index / (level === 'quarter' ? 4 : 12))
  const month = level === 'year' ? 1 : level === 'quarter' ? (index % 4) * 3 + 1 : index % 12 + 1
  return calendarSerial(year, month, 1, date1904)
}

function calendarSerial(year: number, month: number, day: number, date1904: boolean): number {
  const date = parseTimelineDate(`${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`, date1904)
  return date?.serial ?? 0
}

function timelineOffsetAtSerial(bounds: TimelineBounds, serial: number, level: TimelineLevel, date1904: boolean): number {
  const value = clamp(serial, bounds.first.serial, bounds.last.serial)
  const date = parseTimelineDate(value, date1904)
  if (!date) return 0
  const firstIndex = levelIndex(bounds.first, level)
  const index = levelIndex(date, level)
  const start = index === firstIndex ? bounds.first.serial : periodSerial(index, level, date1904)
  const end = periodSerial(index + 1, level, date1904)
  return index - firstIndex + (end > start ? clamp((value - start) / (end - start), 0, 1) : 0)
}
export function timelineSerialAtOffset(items: readonly string[], level: TimelineLevel, date1904: boolean, offset: number): number | undefined {
  const bounds = timelineBounds(items, date1904)
  if (!bounds) return undefined
  const index = levelIndex(bounds.first, level)
  const value = Math.max(0, offset)
  if (value <= 1) return clamp(bounds.first.serial + (periodSerial(index + 1, level, date1904) - bounds.first.serial) * value, bounds.first.serial, bounds.last.serial)
  const whole = Math.floor(value)
  const start = periodSerial(index + whole, level, date1904)
  const end = periodSerial(index + whole + 1, level, date1904)
  return clamp(start + (end - start) * (value - whole), bounds.first.serial, bounds.last.serial)
}
function calendarDate(date: TimelineDate): Date { const result = new Date(0); result.setUTCFullYear(date.year, date.month - 1, date.year === 1900 && date.month === 2 && date.day === 29 ? 28 : date.day); result.setUTCHours(0, 0, 0, 0); return result }
function clamp(value: number, min: number, max: number): number { return Math.min(max, Math.max(min, value)) }
function levelIndex(date: { year: number; month: number; day: number; serial: number }, level: TimelineLevel): number { return level === 'year' ? date.year : level === 'quarter' ? date.year * 4 + Math.floor((date.month - 1) / 3) : level === 'month' ? date.year * 12 + date.month - 1 : Math.floor(date.serial) }
function levelLabel(level: TimelineLevel): string { return level === 'year' ? '年' : level === 'quarter' ? '季度' : level === 'month' ? '月' : '日' }

export function slicerPalette(style: string | undefined, timeline: boolean): { background: string; border: string; header: string; headerBorder: string; headerText: string; item: string; itemBorder: string; selected: string; selectedText: string; text: string; selectionText: string; scrollTrack: string; scrollThumb: string; periodAlternate: string; periodText: string } {
  const match = /Style(Light|Dark)([1-6])$/.exec(style ?? '')
  const accents = ['#3F83F8', '#FF5A1F', '#4E5565', '#F1B312', '#0694A2', '#0DA471']
  const accent = style?.endsWith('Other1') ? '#4E5565' : style?.endsWith('Other2') ? '#1F2124' : accents[Number(match?.[2] ?? 1) - 1] ?? accents[0]!
  const dark = match?.[1] === 'Dark' || style?.endsWith('Other2')
  const background = '#FFFFFF'
  const text = '#4E5565'
  const item = mix(accent, '#FFFFFF', dark ? .76 : .9)
  const selectedText = contrast(accent, '#FFFFFF') >= contrast(accent, '#1B1C1F') ? '#FFFFFF' : '#1B1C1F'
  const selectionText = mix(accent, '#1B1C1F', .42)
  return {
    background, text,
    border: dark ? mix(accent, '#1B1C1F', .5) : accent,
    header: dark ? mix(accent, '#1B1C1F', .24) : '#FFFFFF',
    headerBorder: dark ? mix(accent, '#1B1C1F', .24) : accent,
    headerText: dark ? contrast(mix(accent, '#1B1C1F', .24), '#FFFFFF') >= contrast(mix(accent, '#1B1C1F', .24), '#1B1C1F') ? '#FFFFFF' : '#1B1C1F' : '#1B1C1F',
    item: timeline ? mix(item, accent, .12) : item,
    itemBorder: mix(accent, '#1B1C1F', .5),
    selected: accent,
    selectedText, selectionText,
    scrollTrack: '#FFFFFF', scrollThumb: '#A8B0BD',
    periodAlternate: mix(accent, '#FFFFFF', .94), periodText: '#4E5565',
  }
}

function mix(foreground: string, background: string, ratio: number): string {
  const a = foreground.slice(1).match(/.{2}/g)!.map((part) => parseInt(part, 16))
  const b = background.slice(1).match(/.{2}/g)!.map((part) => parseInt(part, 16))
  return `#${a.map((value, index) => Math.round(value * (1 - ratio) + b[index]! * ratio).toString(16).padStart(2, '0')).join('')}`
}
function contrast(color: string, against: string): number {
  const luminance = (hex: string) => hex.slice(1).match(/.{2}/g)!.map((part) => parseInt(part, 16) / 255).map((value) => value <= .04045 ? value / 12.92 : ((value + .055) / 1.055) ** 2.4).reduce((sum, value, index) => sum + value * [0.2126, 0.7152, 0.0722][index]!, 0)
  const values = [luminance(color), luminance(against)].sort((a, b) => b - a)
  return (values[0]! + .05) / (values[1]! + .05)
}
function contains(rect: RenderRect, x: number, y: number): boolean { return x >= rect.x && y >= rect.y && x <= rect.x + rect.width && y <= rect.y + rect.height }
