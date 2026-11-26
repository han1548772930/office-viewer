import type { SlicerModel, TimelineLevel } from '../../domain/model'
import type { DrawCommand, RenderRect } from '../../rendering/contracts'
import { drawingAffineBounds, drawingAffinePoint, invertDrawingAffine, type DrawingAffineGeometry, type DrawingAffineMatrix } from '../../domain/drawing-affine.ts'
import { intersectRect } from '../../rendering/geometry.ts'
import { GRID_HEADER_HEIGHT, GRID_HEADER_WIDTH, gridViewportHeaders, SpreadsheetGeometry, type GridViewport } from '../../rendering/spreadsheet/layout'
import { parseTimelineDate, timelineDateToInput, type TimelineDate } from './timeline-date'
import { slicerCanvasIcon, type SlicerIconName } from './slicer-canvas-icons'
import { resolveSlicerBounds, slicerDrawingGeometry } from './slicer-drawing'
import type { SlicerControlItemState } from './slicer-control-selection.ts'
import type { SlicerStyleOwner } from './slicer-style-catalog.ts'
import { mixSlicerColor as mix, resolveSlicerRenderProfile, type SlicerCanvasTheme, type SlicerRenderProfile, type SlicerRenderRegion } from './slicer-style-rendering.ts'
import { slicerRegionBorderCommands, slicerRegionTextCommands } from './slicer-style-canvas.ts'
import { effectiveTimelineDateRange, timelineConditionIntersectsPeriod } from './timeline-date-filter.ts'
import { timelinePeriodRange } from './timeline-periods.ts'
import { slicerMemberViewport } from './slicer-member-viewport.ts'
import type { SlicerCanvasDataState, SlicerCanvasMemberWindow } from './slicer-canvas-data.ts'
export { slicerPalette } from './slicer-style-rendering.ts'
export type { SlicerCanvasTheme } from './slicer-style-rendering.ts'

export interface SlicerCanvasControl {
  model: SlicerModel
  items: string[]
  itemState?: SlicerControlItemState
  members?: readonly { label: string; hasData: boolean }[]
  styleOwner?: SlicerStyleOwner
  memberWindow?: SlicerCanvasMemberWindow
  dataStatus?: SlicerCanvasDataState
  /** Inclusive typed source bounds. An explicitly empty value means there is no reliable typed Date source. */
  timeline?: { minDate?: number; maxDate?: number; interactive?: boolean }
}
export interface SlicerCanvasPeriod { key: number; startDate: number; endDate: number; label: string; groupLabel: string }
export interface SlicerObjectLayout {
  id: string
  model: SlicerModel
  styleOwner?: SlicerStyleOwner
  timeline?: { minDate?: number; maxDate?: number; interactive?: boolean }
  dataStatus?: SlicerCanvasDataState
  totalItemCount: number
  /** Content rectangles use a shared local layout; this matrix maps them to header-inclusive world pixels. */
  matrix: DrawingAffineMatrix
  inverseMatrix: DrawingAffineMatrix
  geometry: DrawingAffineGeometry
  bounds: RenderRect
  rect: RenderRect
  header: RenderRect
  clear: RenderRect
  multiSelect: RenderRect
  items: Array<{ value: string; rect: RenderRect; selected: boolean; partial?: boolean; hasData?: boolean; index: number }>
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

export function slicerObjectLayouts(geometry: SpreadsheetGeometry, controls: readonly SlicerCanvasControl[], date1904: boolean, locale = 'zh-CN',
  previews?: ReadonlyMap<string, DrawingAffineGeometry>, scrollPreviews?: ReadonlyMap<string, number>): SlicerObjectLayout[] {
  return controls.map((control, index) => ({ ...control, order: control.model.drawing?.zIndex ?? index }))
    .sort((first, second) => first.order - second.order).flatMap<SlicerObjectLayout>(({ model, items, itemState, members, styleOwner, timeline, memberWindow, dataStatus }) => {
    const anchored = resolveSlicerBounds(model, geometry)
    if (anchored.width <= 0 || anchored.height <= 0) return []
    const drawing = previews?.get(`${model.type}:${model.id}`) ?? slicerDrawingGeometry(model, geometry),
      bounds = drawingAffineBounds(drawing), x = bounds.left + GRID_HEADER_WIDTH, y = bounds.top + GRID_HEADER_HEIGHT,
      { width, height, matrix: m } = drawing,
      matrix = { ...m, e: m.e + GRID_HEADER_WIDTH - m.a * x - m.c * y, f: m.f + GRID_HEADER_HEIGHT - m.b * x - m.d * y },
      projection = { geometry: drawing, matrix, inverseMatrix: invertDrawingAffine(matrix), bounds: { x, y, width: bounds.width, height: bounds.height }, styleOwner, timeline, dataStatus }
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
    const scrollPreview = scrollPreviews?.get(`${model.type}:${model.id}`)
    if (model.type === 'timeline') return [{ ...timelineLayout(model, items, rect, header, clear, level, date1904, locale, scrollPreview, timeline), ...projection }]

    const rowHeight = model.rowHeight ?? 24
    const totalItemCount = memberWindow?.total ?? items.length
    const viewport = slicerMemberViewport(model, width, height, totalItemCount, scrollPreview ?? model.startItem ?? 0)
    const columnCount = viewport.columns
    const contentY = headerHeight
    const visibleRows = viewport.rows, pitch = viewport.pitch, totalRows = viewport.totalRows,
      maxScrollRows = viewport.maxScrollOffset / columnCount, scrollRow = viewport.start / columnCount,
      firstItem = viewport.start, scrollWidth = viewport.scrollWidth, itemWidth = viewport.itemWidth
    const visibleItems = memberWindow
      ? memberWindow.items.map((member, index) => ({ value: member.label, index: memberWindow.start + index, member }))
        .filter(item => item.index >= firstItem && item.index < firstItem + viewport.count)
      : items.slice(firstItem, Math.min(items.length, firstItem + viewport.count)).map((value, index) => ({ value, index: firstItem + index, member: undefined }))
    const checked = itemState ? new Set(itemState.checkedItems) : undefined
    const partial = new Set(itemState?.partialItems)
    const availability = new Map<string, boolean>()
    for (const member of members ?? []) availability.set(member.label, availability.get(member.label) === true || member.hasData)
    const layoutItems = visibleItems.map(({ value, index, member }) => {
      const visibleIndex = index - firstItem
      const row = Math.floor(visibleIndex / columnCount)
      const column = visibleIndex % columnCount
      const filterItems = model.selectedItems
      return {
        value,
        rect: { x: x + 4 + column * (itemWidth + 4), y: y + contentY + 4 + row * pitch, width: itemWidth, height: rowHeight },
        selected: member?.selected ?? (checked ? checked.has(value) : filterItems == null || filterItems.includes(value)),
        partial: member?.partial ?? partial.has(value),
        hasData: member?.hasData ?? availability.get(value) ?? true,
        index,
      }
    })
    const scrollbar = scrollWidth ? { x: x + width - 14, y: y + contentY + 4, width: 10, height: Math.max(0, height - contentY - 8) } : undefined
    const thumbHeight = scrollbar ? Math.min(scrollbar.height, Math.max(12, scrollbar.height * visibleRows / totalRows)) : 0
    const scrollThumb = scrollbar ? { ...scrollbar, y: scrollbar.y + (maxScrollRows ? (scrollbar.height - thumbHeight) * scrollRow / maxScrollRows : 0), height: thumbHeight } : undefined
    return [{
      id: model.id, model, rect, header, clear, multiSelect, level,
      items: layoutItems, periods: [], scrollbar, scrollThumb, totalItemCount,
      content: { x, y: y + contentY, width, height: Math.max(0, height - contentY) },
      scrollOffset: firstItem, maxScrollOffset: maxScrollRows * columnCount, visibleCount: visibleRows * columnCount,
      ...projection,
    }]
  })
}

function timelineLayout(model: SlicerModel, items: readonly string[], rect: RenderRect, header: RenderRect, clear: RenderRect, level: RenderRect, date1904: boolean, locale: string, scrollPreview?: number,
  source?: SlicerCanvasControl['timeline']): Omit<SlicerObjectLayout, 'matrix' | 'inverseMatrix' | 'geometry' | 'bounds'> {
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
  const bounds = timelineBounds(items, date1904, source)
  const currentLevel = model.level ?? 'month'
  const total = bounds ? levelIndex(bounds.last, currentLevel) - levelIndex(bounds.first, currentLevel) + 1 : 0
  const maxScrollOffset = Math.max(0, total - visibleCount)
  const first = clamp(scrollPreview ?? (bounds ? timelineOffsetAtSerial(bounds, model.scrollPosition ?? bounds.first.serial, currentLevel, date1904) : 0), 0, maxScrollOffset)
  const firstVisible = Math.max(0, Math.floor(first) - 1)
  const shown = bounds ? periodWindow(bounds, currentLevel, date1904, firstVisible, visibleCount + 2, locale) : []
  const selection = model.timelineDateRange === undefined ? undefined : effectiveTimelineDateRange(model.timelineDateRange)
  const periodLayouts = shown.map((period, index) => {
    const absoluteIndex = firstVisible + index
    const slotX = x + absoluteIndex * TIMELINE_SLOT - first * TIMELINE_SLOT
    return {
      period,
      rect: { x: slotX, y: bandY, width: TIMELINE_SLOT, height: groupHeight + labelHeight + trackHeight },
      selected: selection !== undefined ? period.endDate > selection.startDate && period.startDate < selection.endDate
        : model.timelineDateCondition !== undefined && timelineConditionIntersectsPeriod(model.timelineDateCondition, period),
    }
  })
  const scrollTrack = scrollHeight ? { x: x + 5, y: y + height - scrollHeight, width: Math.max(0, width - 10), height: scrollHeight } : undefined
  const thumbWidth = scrollTrack ? Math.min(scrollTrack.width, Math.max(16, scrollTrack.width * visibleCount / Math.max(visibleCount, total))) : 0
  const scrollThumb = scrollTrack ? { ...scrollTrack, x: scrollTrack.x + (total > visibleCount ? (scrollTrack.width - thumbWidth) * first / (total - visibleCount) : 0), width: thumbWidth } : undefined
  const menuRowHeight = Math.min(24, Math.max(0, y + height - level.y - level.height) / 4)
  const menu = TIMELINE_LEVELS.map((_, index) => ({ x: level.x, y: level.y + level.height + index * menuRowHeight, width: level.width, height: menuRowHeight }))
  const rangeLabel = selection ? `${timelineDateToInput(selection.startDate, date1904)} – ${timelineDateToInput(selection.endDate, date1904)}（不含）`
    : model.timelineDateCondition ? '日期条件筛选' : model.selectedItems !== undefined ? '保留的成员筛选' : '所有时段'
  const selected = periodLayouts.filter(period => period.selected)
  const handleWidth = Math.min(8, width)
  const boundaryX = (serial: number, entry: typeof selected[number]) => entry.rect.x
    + clamp((serial - entry.period.startDate) / (entry.period.endDate - entry.period.startDate), 0, 1) * TIMELINE_SLOT
  const startHandle = selection && selected.length ? { x: clamp(boundaryX(selection.startDate, selected[0]!) - handleWidth / 2, x, x + width - handleWidth), y: axisTrack.y, width: handleWidth, height: axisTrack.height } : undefined
  const endHandle = selection && selected.length ? { x: clamp(boundaryX(selection.endDate, selected[selected.length - 1]!) - handleWidth / 2, x, x + width - handleWidth), y: axisTrack.y, width: handleWidth, height: axisTrack.height } : undefined
  return { id: model.id, model, rect, header, clear, multiSelect: { x, y, width: 0, height: 0 }, items: [], periods: periodLayouts, totalItemCount: 0, track: axisTrack, groupLabelBand, periodLabelBand, level, scrollbar: scrollTrack, scrollThumb, selectionLabel: model.showSelectionLabel === false ? undefined : rangeLabel, selectionLabelRect: { x, y: y + header.height, width: Math.max(0, width - (model.showTimeLevel === false ? 0 : level.width)), height: secondaryHeight }, menu, startHandle, endHandle, scrollOffset: first, maxScrollOffset, visibleCount }
}

export interface SlicerDrawState {
  timelineMenuId?: string
  selectedSlicerId?: string
  multiSelect?: ReadonlySet<string>
  focusedItems?: ReadonlyMap<string, number>
  focusedPeriods?: ReadonlyMap<string, number>
  onIconReady?: () => void
  /** Interactive workbenches mount Ant controls; print/export can draw the same values without DOM. */
  controlUI?: 'canvas' | 'antd'
  theme?: SlicerCanvasTheme
  styleProfiles?: ReadonlyMap<string, SlicerRenderProfile>
  hoveredItems?: ReadonlyMap<string, number>
}

export function slicerObjectDrawCommands(layouts: readonly SlicerObjectLayout[], state: SlicerDrawState = {}): DrawCommand[] {
  const commands: DrawCommand[] = []
  for (const layout of layouts) {
    const profile = state.styleProfiles?.get(`${layout.model.type}:${layout.id}`)
      ?? resolveSlicerRenderProfile(layout.styleOwner ?? {}, layout.model.type, layout.model.style, state.theme)
    const colors = profile.palette, antControls = state.controlUI === 'antd'
    const begin = commands.length
    commands.push({ kind: 'rect', ...layout.rect, fill: profile.whole.fill })
    if (layout.header.height > 0) {
      commands.push({ kind: 'rect', ...layout.header, fill: profile.header.fill }, ...slicerRegionBorderCommands(layout.header, profile.header))
      commands.push(...slicerRegionTextCommands({ ...layout.header, width: Math.max(0, layout.header.width - 52) }, layout.model.caption ?? layout.model.name ?? '', profile.header))
      const filtered = layout.model.type === 'timeline'
        ? layout.model.timelineDateRange !== undefined || layout.model.timelineDateCondition !== undefined || layout.model.timelineRange !== undefined || layout.model.selectedItems !== undefined
        : layout.model.selectedItems !== undefined
      if (!antControls) drawIcon(commands, layout.clear, 'clear', colors.headerText, state.onIconReady, filtered ? 1 : .48)
      if (!antControls && layout.model.type === 'slicer') {
        const active = state.multiSelect?.has(layout.id)
        if (active) roundedRect(commands, layout.multiSelect, 6, mix(colors.background, colors.selected, .16))
        drawIcon(commands, layout.multiSelect, 'multi', colors.headerText, state.onIconReady, active ? 1 : .48)
      }
    }
    if (layout.model.type === 'slicer') {
      for (const item of antControls ? [] : layout.items) {
        const focused = state.focusedItems?.get(layout.id) === item.index
        const style = profile.type === 'slicer' ? profile.items[
          state.hoveredItems?.get(layout.id) === item.index ? item.selected ? item.hasData === false ? 'hoveredSelectedItemWithNoData' : 'hoveredSelectedItemWithData'
            : item.hasData === false ? 'hoveredUnselectedItemWithNoData' : 'hoveredUnselectedItemWithData'
            : item.selected ? item.hasData === false ? 'selectedItemWithNoData' : 'selectedItemWithData'
              : item.hasData === false ? 'unselectedItemWithNoData' : 'unselectedItemWithData'] : profile.whole
        roundedRect(commands, item.rect, 6, item.partial ? mix(style.fill, colors.selected, .24) : style.fill)
        commands.push(...slicerRegionBorderCommands(item.rect, style), ...slicerRegionTextCommands(item.rect, item.value, style))
        if (focused) commands.push({ kind: 'rect', ...item.rect, fill: 'transparent', stroke: colors.selected, strokeWidth: 1.25 })
        if (item.partial) commands.push({ kind: 'line', x1: item.rect.x + Math.max(0, item.rect.width - 16), y1: item.rect.y + item.rect.height - 5,
          x2: item.rect.x + Math.max(0, item.rect.width - 7), y2: item.rect.y + item.rect.height - 5, color: colors.selected, width: 2, clip: layout.content })
      }
    } else if (profile.type === 'timeline') {
      if (layout.selectionLabel && layout.selectionLabelRect) {
        appendTimelineRegion(commands, layout.selectionLabelRect, profile.elements.selectionLabel, layout.selectionLabel, 12, 5)
      }
      if (layout.model.showTimeLevel !== false) {
        appendTimelineRegion(commands, layout.level, profile.elements.timeLevel)
        if (!antControls) {
          commands.push(...slicerRegionTextCommands({ ...layout.level, width: Math.max(0, layout.level.width - 18) }, levelLabel(layout.model.level ?? 'month'), profile.elements.timeLevel, 12, 5))
          drawIcon(commands, { x: layout.level.x + layout.level.width - 17, y: layout.level.y + (layout.level.height - 16) / 2, width: 16, height: 16 }, state.timelineMenuId === layout.id ? 'moreUp' : 'moreDown', profile.elements.timeLevel.color, state.onIconReady)
        }
      }
      const group = layout.groupLabelBand
      const label = layout.periodLabelBand
      const track = layout.track
      if (group && label && track) {
        const clip = { x: layout.rect.x, y: group.y, width: layout.rect.width, height: group.height + label.height + track.height }
        for (let index = 0; index < layout.periods.length; index += 1) {
          const period = layout.periods[index]!
          const px = period.rect.x
          const periodStyle = profile.elements[index % 2 === 0 ? 'periodLabel1' : 'periodLabel2']
          appendTimelineRegion(commands, { x: px, y: label.y, width: period.rect.width, height: label.height }, periodStyle, period.period.label, 12, 5, clip)
          if (!antControls) {
            appendTimelineRegion(commands, { x: px, y: track.y, width: period.rect.width, height: track.height },
              profile.elements[period.selected ? 'selectedTimeBlock' : 'unselectedTimeBlock'], undefined, 12, 5, clip)
            appendTimelineRegion(commands, { x: px + period.rect.width - 1, y: track.y, width: 1, height: track.height },
              period.selected ? profile.elements.selectedTimeBlockSpace : profile.whole, undefined, 12, 0, clip)
          }
          if (period.period.groupLabel && (!index || layout.periods[index - 1]!.period.groupLabel !== period.period.groupLabel)) {
            const nextGroup = layout.periods.slice(index + 1).find(next => next.period.groupLabel !== period.period.groupLabel)
            appendTimelineRegion(commands, { x: px, y: group.y, width: Math.max(period.rect.width, (nextGroup?.rect.x ?? layout.rect.x + layout.rect.width) - px), height: group.height },
              periodStyle, period.period.groupLabel, 10, 5, clip)
          }
          if (!antControls && state.focusedPeriods?.get(layout.id) === period.period.key) commands.push({ kind: 'rect', x: px + 1, y: label.y + 1, width: Math.max(0, period.rect.width - 2), height: Math.max(0, label.height + track.height - 2), fill: 'transparent', stroke: colors.border, strokeWidth: 1, clip })
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
    if (!antControls && layout.scrollbar && layout.scrollThumb) {
      commands.push({ kind: 'rect', ...layout.scrollbar, fill: colors.background, stroke: colors.border, strokeWidth: 1 })
      if (layout.model.type === 'timeline') {
        roundedRect(commands, { x: layout.scrollThumb.x + 1, y: layout.scrollThumb.y + Math.max(1, (layout.scrollThumb.height - 4) / 2), width: Math.max(0, layout.scrollThumb.width - 2), height: Math.min(4, Math.max(0, layout.scrollThumb.height - 2)) }, 2, colors.border)
      } else commands.push({ kind: 'rect', x: layout.scrollThumb.x + 1, y: layout.scrollThumb.y + 1, width: Math.max(0, layout.scrollThumb.width - 2), height: Math.max(0, layout.scrollThumb.height - 2), fill: colors.border })
    }
    if (!antControls && state.timelineMenuId === layout.id) layout.menu?.forEach((menuItem, index) => {
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
    commands.push(...slicerRegionBorderCommands(layout.rect, profile.whole))
    for (let index = begin; index < commands.length; index += 1) {
      const command = commands[index]!
      command.localClip = command.clip ? intersectRect(command.clip, layout.rect) : { ...layout.rect }
      delete command.clip
      command.matrix = { ...layout.matrix }
      command.drawingRef = { kind: layout.model.type === 'timeline' ? 'timeline' : 'slicer', id: layout.id }
    }
  }
  return commands
}

function appendTimelineRegion(commands: DrawCommand[], rect: RenderRect, style: SlicerRenderRegion, text?: string, fontSize = 12, padding = 5, clip?: RenderRect): void {
  if (rect.width <= 0 || rect.height <= 0) return
  const projected: DrawCommand[] = [{ kind: 'rect', ...rect, fill: style.fill }, ...slicerRegionBorderCommands(rect, style),
    ...(text === undefined ? [] : slicerRegionTextCommands(rect, text, style, fontSize, padding))]
  for (const command of projected) {
    if (clip) command.clip = command.clip ? intersectRect(command.clip, clip) : clip
    commands.push(command)
  }
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
    const local = slicerObjectLocalPoint(layout, { x, y }), hit = hitTestSlicerLocalObject(layout, local.x, local.y, timelineMenuId)
    if (hit) return hit
  }
  return undefined
}

export function slicerObjectLocalPoint(layout: SlicerObjectLayout, point: { x: number; y: number }): { x: number; y: number } {
  return drawingAffinePoint(layout.inverseMatrix, point.x, point.y)
}

function hitTestSlicerLocalObject(layout: SlicerObjectLayout, x: number, y: number, timelineMenuId: string): SlicerCanvasHit | undefined {
    if (!contains(layout.rect, x, y)) return undefined
    if (timelineMenuId === layout.id && layout.menu?.some((rect) => contains(rect, x, y))) return { type: 'timelineMenu', id: layout.id }
    if (layout.header.height && contains(layout.clear, x, y)) return { type: 'clear', id: layout.id }
    if (layout.model.type === 'slicer' && layout.header.height && contains(layout.multiSelect, x, y)) return { type: 'multiSelect', id: layout.id }
    if (layout.model.type === 'timeline') {
      if (layout.model.showTimeLevel !== false && contains(layout.level, x, y)) return { type: 'timelineLevel', id: layout.id }
      if (layout.timeline?.interactive !== false && layout.startHandle && contains(layout.startHandle, x, y)) return { type: 'timelineStart', id: layout.id }
      if (layout.timeline?.interactive !== false && layout.endHandle && contains(layout.endHandle, x, y)) return { type: 'timelineEnd', id: layout.id }
    }
    const itemIndex = layout.items.findIndex((item) => contains(item.rect, x, y))
    if (itemIndex >= 0) return { type: 'item', id: layout.id, value: layout.items[itemIndex]!.value, index: layout.items[itemIndex]!.index }
    const periodIndex = layout.periods.findIndex((period) => contains(period.rect, x, y))
    if (periodIndex >= 0 && layout.timeline?.interactive !== false) return { type: 'period', id: layout.id, period: layout.periods[periodIndex]!.period, index: layout.periods[periodIndex]!.period.key }
    if (layout.scrollbar && contains(layout.scrollbar, x, y)) return { type: 'scroll', id: layout.id, part: layout.scrollThumb && contains(layout.scrollThumb, x, y) ? 'thumb' : 'track' }
    if (!layout.model.lockedPosition && contains({ x: layout.rect.x + layout.rect.width - 10, y: layout.rect.y + layout.rect.height - 10, width: 10, height: 10 }, x, y)) return { type: 'resize', id: layout.id }
    if (!layout.model.lockedPosition && layout.header.height > 0 && y <= layout.header.y + layout.header.height) return { type: 'move', id: layout.id }
    return { type: 'body', id: layout.id }
}

export function slicerObjectContentPoint(x: number, y: number, viewport: GridViewport): { x: number; y: number } {
  const headers = gridViewportHeaders(viewport)
  return { x: GRID_HEADER_WIDTH + (x - headers.width + viewport.scrollLeft) / viewport.zoom,
    y: GRID_HEADER_HEIGHT + (y - headers.height + viewport.scrollTop) / viewport.zoom }
}

export function timelinePeriods(items: readonly string[], level: TimelineLevel, date1904: boolean, startIndex = 0, count = Number.POSITIVE_INFINITY, locale = 'zh-CN'): SlicerCanvasPeriod[] {
  const bounds = timelineBounds(items, date1904)
  return bounds ? periodWindow(bounds, level, date1904, startIndex, count, locale) : []
}

interface TimelineBounds { first: TimelineDate; last: TimelineDate }
function timelineBounds(items: readonly string[], date1904: boolean, source?: SlicerCanvasControl['timeline']): TimelineBounds | undefined {
  if (source !== undefined) {
    const first = source.minDate === undefined ? undefined : parseTimelineDate(source.minDate, date1904)
    const last = source.maxDate === undefined ? undefined : parseTimelineDate(source.maxDate, date1904)
    return first && last && first.serial <= last.serial ? { first, last } : undefined
  }
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
    const { endDate } = timelinePeriodRange(serial, level, date1904)
    const groupLabel = level === 'day' ? new Intl.DateTimeFormat(locale, { month: 'short', year: 'numeric', timeZone: 'UTC' }).format(calendarDate(date)) : level === 'year' ? '' : String(date.year)
    periods.push({ key, startDate: serial, endDate, label, groupLabel })
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
  const end = timelinePeriodRange(date.serial, level, date1904).endDate
  return index - firstIndex + (end > start ? clamp((value - start) / (end - start), 0, 1) : 0)
}
export function timelineSerialAtOffset(items: readonly string[], level: TimelineLevel, date1904: boolean, offset: number, source?: SlicerCanvasControl['timeline']): number | undefined {
  const bounds = timelineBounds(items, date1904, source)
  if (!bounds) return undefined
  const index = levelIndex(bounds.first, level)
  const value = Math.max(0, offset)
  if (value <= 1) return clamp(bounds.first.serial + (timelinePeriodRange(bounds.first.serial, level, date1904).endDate - bounds.first.serial) * value, bounds.first.serial, bounds.last.serial)
  const whole = Math.floor(value)
  const start = periodSerial(index + whole, level, date1904)
  const date = parseTimelineDate(start, date1904)
  if (!date) return bounds.last.serial
  const end = timelinePeriodRange(start, level, date1904).endDate
  return clamp(start + (end - start) * (value - whole), bounds.first.serial, bounds.last.serial)
}
function calendarDate(date: TimelineDate): Date { const result = new Date(0); result.setUTCFullYear(date.year, date.month - 1, date.year === 1900 && date.month === 2 && date.day === 29 ? 28 : date.day); result.setUTCHours(0, 0, 0, 0); return result }
function clamp(value: number, min: number, max: number): number { return Math.min(max, Math.max(min, value)) }
function levelIndex(date: { year: number; month: number; day: number; serial: number }, level: TimelineLevel): number { return level === 'year' ? date.year : level === 'quarter' ? date.year * 4 + Math.floor((date.month - 1) / 3) : level === 'month' ? date.year * 12 + date.month - 1 : Math.floor(date.serial) }
function levelLabel(level: TimelineLevel): string { return level === 'year' ? '年' : level === 'quarter' ? '季度' : level === 'month' ? '月' : '日' }

function contains(rect: RenderRect, x: number, y: number): boolean { return x >= rect.x && y >= rect.y && x <= rect.x + rect.width && y <= rect.y + rect.height }
