import { createSlicer, setSlicerRecordSelection } from '../../src/features/slicer/slicer-engine.ts'
import { slicerRecordSelectionWorkbook, slicerTableRecordSelectionWorkbook } from './pivot-filter-fixture.mjs'

/** Synthetic public-format style definitions attached to the existing seven-record native/report sample. */
export const slicerStyleStates = ['unselectedItemWithData', 'selectedItemWithData', 'unselectedItemWithNoData', 'selectedItemWithNoData',
  'hoveredUnselectedItemWithData', 'hoveredSelectedItemWithData', 'hoveredUnselectedItemWithNoData', 'hoveredSelectedItemWithNoData']
export const timelineStyleElements = ['selectionLabel', 'timeLevel', 'periodLabel1', 'periodLabel2', 'selectedTimeBlock', 'unselectedTimeBlock', 'selectedTimeBlockSpace']
export const styleFormatting = {
  bg: { rgb: '#102030', th: 4 }, cl: { rgb: null, th: 1 }, bl: 1, ht: 5, vt: 2,
  ol: { s: 1, c: 0, t: 2, cl: { rgb: '#abcdef', th: 3 } },
  bd: { t: { s: 1, cl: { rgb: '#123456', th: 5 } }, r: null, b: { s: 13, cl: {} }, l: null,
    tl_br: { s: 2, cl: { th: 11 } }, tl_bc: null, tl_mr: null, bl_tr: null, ml_tr: null, bc_tr: null },
}
export function customSlicerStyle(name = 'Regional members') {
  return { name, wholeStyle: structuredClone(styleFormatting), headerRowStyle: { bg: null, cl: { rgb: '#fedcba' }, bl: 0, ht: null, vt: null, bd: null },
    slicerStyleElements: slicerStyleStates.map((type, index) => ({ type,
      formatting: { bg: { th: index }, cl: index === 2 ? null : { rgb: '#334455' }, bl: index % 2, ht: 2, vt: 2 } })) }
}
export function customTimelineStyle(name = 'Calendar bands') {
  return { name, wholeStyle: structuredClone(styleFormatting), headerRowStyle: { bg: null, bl: 0 },
    timelineStyleElements: timelineStyleElements.map((type, index) => ({ type,
      formatting: { bg: { th: index + 1 }, cl: { rgb: '#556677' }, bl: index % 2 } })) }
}
export function slicerStyleWorkbook(kind = 'pivot') {
  const workbook = kind === 'table' ? slicerTableRecordSelectionWorkbook() : slicerRecordSelectionWorkbook()
  const [defaultView, explicitView] = workbook.slicers
  setSlicerRecordSelection(workbook, defaultView, { mode: 'include', values: [1] })
  const timeline = createSlicer(workbook, { id: 'style-timeline', unitId: workbook.id, hostSheetId: 'dashboard',
    pivotTableId: workbook.pivotTables[0].id, sourceFieldId: 'date-source', type: 'timeline' })
  delete defaultView.style; explicitView.style = 'SlicerStyleDark2'; delete timeline.style
  workbook.slicerStyles = { defaultSlicerStyle: 'Regional members', styles: [customSlicerStyle(), { name: 'Unused slicer', slicerStyleElements: [] }] }
  workbook.timelineStyles = { defaultTimelineStyle: 'Calendar bands', styles: [customTimelineStyle(), { name: 'Unused timeline', timelineStyleElements: [] }] }
  workbook.metadata = { sourceFixture: 'seven native/report records', preservedField: { author: 'fixture' } }
  return workbook
}
