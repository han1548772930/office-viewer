import { slicerConnectionsWorkbook, connectionFixtureIds as ids } from './slicer-connections-fixture.mjs'
import { patchSlicerCacheOptions } from '../../src/features/slicer/slicer-cache-options.ts'
import { setSlicerTimelineRange, updateSlicerPresentation } from '../../src/features/slicer/slicer-engine.ts'

// Real source/cache/control/report fixture. No app or checks have run for this draft.
export function slicerSettingsUiWorkbook(options = {}) {
  const workbook = slicerConnectionsWorkbook(options)
  if (options.customStyles) workbook.slicerStyles = { defaultSlicerStyle: 'Office fixture style', styles: [{ name: 'Office fixture style',
    wholeStyle: { bg: { rgb: '#FFFFFF' }, cl: { rgb: '#123456' } }, headerRowStyle: { bg: { rgb: '#103F77' }, cl: { rgb: '#FFFFFF' } },
    slicerStyleElements: [{ type: 'selectedItemWithData', formatting: { bg: { rgb: '#103F77' }, cl: { rgb: '#FFFFFF' } } }] }] }
  for (const control of workbook.slicers.filter(control => control.type === 'slicer')) {
    if (options.viewDefaults) updateSlicerPresentation(control, { caption: undefined, startItem: undefined, columnCount: undefined,
      showCaption: undefined, lockedPosition: undefined, style: undefined, rowHeight: 24 }, workbook)
    if (options.captionInherited) updateSlicerPresentation(control, { caption: undefined }, workbook)
  }
  const primary = workbook.slicers.find(control => control.id === ids.slicer)
  if (options.authoredSettings) updateSlicerPresentation(primary, { name: 'Region view', caption: 'Region caption', startItem: 1,
    rowHeight: 31.125, columnCount: 2, showCaption: false, lockedPosition: false, style: 'SlicerStyleDark2' }, workbook)
  if (options.customStyles && options.customOverride) updateSlicerPresentation(primary, { style: 'Office fixture style' }, workbook)
  if (options.unknownStyle) primary.style = 'Missing imported style'
  if (options.cacheOptions) patchSlicerCacheOptions(workbook, primary, options.cacheOptions)
  if (options.marchContext) setSlicerTimelineRange(workbook, workbook.slicers.find(control => control.id === ids.timeline),
    { start: '2026-03-01', end: '2026-03-31', level: 'month' })
  if (options.hiddenSource) workbook.sheets.find(sheet => sheet.id === ids.source).worksheetProtection = {
    id: 'hidden-settings-source', ownerUserId: 'owner', allowedUserIds: [], editState: 'onlyMe', viewState: 'noOneElseCanView',
  }
  return workbook
}
