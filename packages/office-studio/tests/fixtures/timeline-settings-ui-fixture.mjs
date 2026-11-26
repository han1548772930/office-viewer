import { createCell } from '../../src/domain/model.ts'
import { updateSlicerPresentation } from '../../src/features/slicer/slicer-engine.ts'
import { patchTimelineSettings } from '../../src/features/slicer/timeline-settings.ts'
import { setTimelineDateCondition, setTimelineDateRange } from '../../src/features/slicer/timeline-filter.ts'
import { timelineFilterWorkbook, timelineFirst, timelineViews } from './timeline-filter-fixture.mjs'

export const timelineUiIds = { primary: 'timeline-primary', shared: 'timeline-copy', slicer: 'timeline-mixed-slicer',
  host: 'dashboard', source: 'data', other: 'output' }

/** Real original date records, raw colliders, cached Date, shared views and all reports. No checks have run. */
export function timelineSettingsUiWorkbook(options = {}) {
  const workbook = timelineFilterWorkbook(options), first = timelineFirst(workbook), views = timelineViews(workbook)
  workbook.activeSheetId = timelineUiIds.host
  views.forEach((view, index) => {
    updateSlicerPresentation(view, { layout: { left: 120, top: 200 + index * 220, width: 500, height: 180 } }, workbook)
    if (options.viewDefaults) patchTimelineSettings(workbook, view, { level: 2, selectionLevel: 2, caption: undefined, style: undefined,
      scrollPosition: undefined, showHeader: undefined, showSelectionLabel: undefined, showTimeLevel: undefined, showHorizontalScrollbar: undefined })
    if (options.captionInherited) patchTimelineSettings(workbook, view, { caption: undefined })
  })
  const primary = views[0], source = workbook.sheets.find(sheet => sheet.id === timelineUiIds.source)
  if (options.authoredSettings) patchTimelineSettings(workbook, primary, { name: 'Timeline authored', caption: 'Authored calendar',
    level: 0, selectionLevel: 3, scrollPosition: first + .123456789, style: 'TimelineStyleDark2',
    showHeader: false, showSelectionLabel: false, showTimeLevel: false, showHorizontalScrollbar: false })
  if (options.customOverride) patchTimelineSettings(workbook, primary, { style: 'Calendar bands' })
  if (options.unknownStyle) primary.style = 'Missing imported calendar style'
  if (options.range) setTimelineDateRange(workbook, primary, options.range)
  if (options.condition) setTimelineDateCondition(workbook, primary, options.condition)
  if (options.compatibilityDay) {
    for (let row = 3; row <= 13; row += 1) {
      const cell = createCell(row, 4, 60 + (row - 3) / 24)
      cell.style.numberFormat = 'yyyy-mm-dd hh:mm:ss'; source.cells[`${row}:4`] = cell
    }
    setTimelineDateRange(workbook, primary, { startDate: 60.25, endDate: 60.75 })
    patchTimelineSettings(workbook, primary, { scrollPosition: 60.125 })
  }
  if (options.noDates) for (let row = 3; row <= 13; row += 1) source.cells[`${row}:4`] = createCell(row, 4, row)
  if (options.hiddenSource) source.worksheetProtection = { id: 'timeline-hidden-source', ownerUserId: 'owner', allowedUserIds: [],
    editState: 'onlyMe', viewState: 'noOneElseCanView' }
  if (options.readonly) workbook.sheets.find(sheet => sheet.id === timelineUiIds.host).worksheetProtection = {
    id: 'timeline-ui-readonly', ownerUserId: 'owner', allowedUserIds: [], editState: 'onlyMe', viewState: 'othersCanView' }
  return workbook
}
