import { createCell } from '../../src/domain/model.ts'
import { setAutoFilter, setFilterCriteria } from '../../src/features/sheet-sort-filter/sort-filter-engine.ts'
import { patchSlicerCacheOptions } from '../../src/features/slicer/slicer-cache-options.ts'
import { slicerTableRecordSelectionWorkbook } from './pivot-filter-fixture.mjs'

// Shared deterministic input only. Production algorithms are never replaced.
export const SLICER_WORKER_DATASET_VERSION = 'typed-members-v1-lcg-1664525-1013904223'
export const SLICER_WORKER_SEED = 0x51_1c_e2
export const slicerWorkerDefaults = { mode: 'items', start: 0, count: 100, locale: 'en-US',
  sortOrder: 'ascending', customListSort: false, crossFilter: 'showItemsWithNoData', hideItemsWithNoData: false }

export function fixedSlicerWorkerRows() {
  return [{ type: 'blank' }, { type: 'number', value: 1 }, { type: 'date', value: 1 },
    { type: 'string', value: '1' }, { type: 'boolean', value: true }, { type: 'string', value: 'true' },
    { type: 'error', value: '#N/A' }, { type: 'string', value: '#N/A' }, { type: 'string', value: '' },
    { type: 'number', value: 1 }, { type: 'number', value: -0 }, { type: 'number', value: 0 },
    { type: 'date', value: 5.5 }, { type: 'date', value: -2.25 }]
}
export const fixedSlicerWorkerAvailability = [false, false, false, true, true, false, false, true,
  false, true, false, true, true, true]

export function generateSlicerWorkerRows(rowCount, seed = SLICER_WORKER_SEED) {
  if (!Number.isSafeInteger(rowCount) || rowCount < 1 || rowCount > 1_000_000) throw new RangeError('Invalid fixture row count')
  let state = seed >>> 0
  const rows = [], availability = []
  const vocabulary = ['January', 'Jan', 'February', 'Sunday', 'Sun', 'ALPHA', 'alpha', '', 'a\0b', '😀', '\ud800', '\ud801']
  for (let row = 0; row < rowCount; row++) {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0
    const value = (state >>> 8) % Math.max(8, Math.floor(rowCount / 3))
    const kind = row % 12
    rows.push(kind === 0 ? { type: 'blank' } : kind === 1 || kind === 2 ? { type: 'number', value: value - 100 }
      : kind === 3 || kind === 4 ? { type: 'date', value: value - 2.25 }
      : kind === 5 ? { type: 'boolean', value: ((state >>> 9) & 1) === 1 }
      : kind === 6 ? { type: 'error', value: row % 24 === 6 ? '#N/A' : '#DIV/0!' }
      : { type: 'string', value: kind === 7 ? vocabulary[value % vocabulary.length] : `member${value}` })
    availability.push((state & 7) < 3)
  }
  return { rows, availability, seed, version: SLICER_WORKER_DATASET_VERSION }
}

export function fingerprintSlicerWorkerRows(rows, availability) {
  let checksum = 2166136261
  rows.forEach((raw, index) => {
    const text = `${raw.type}:${raw.type === 'blank' ? '' : String(raw.value)}:${availability[index] ? 1 : 0}\n`
    for (let offset = 0; offset < text.length; offset++) checksum = Math.imul(checksum ^ text.charCodeAt(offset), 16777619) >>> 0
  })
  return checksum.toString(16).padStart(8, '0')
}

/** Real native Table model with an independent Region predicate. The queried
 * Mixed field and its original source-row offsets are retained verbatim.
 */
export function slicerWorkerTableWorkbook(rows, availability, options = {}) {
  if (!rows.length || rows.length !== availability.length) throw new RangeError('Table fixture needs matching nonempty rows')
  const workbook = slicerTableRecordSelectionWorkbook(), slicer = workbook.slicers[0]
  const table = workbook.tables.find(candidate => candidate.id === 'typed-table'), sheet = workbook.sheets.find(candidate => candidate.id === 'data')
  table.range.endRow = rows.length + 2
  sheet.rowCount = Math.max(sheet.rowCount, rows.length + 3)
  rows.forEach((raw, index) => {
    const row = index + 3, cell = createCell(row, 6, raw.type === 'blank' ? null : raw.value)
    if (raw.type === 'date') cell.dataType = 'date'
    if (raw.type === 'error') { cell.dataType = 'e'; cell.metadata = { error: raw.value } }
    sheet.cells[`${row}:6`] = cell
    sheet.cells[`${row}:3`] = createCell(row, 3, availability[index] ? 'North' : 'South')
  })
  const filterSheet = { ...sheet, autoFilter: undefined }
  setAutoFilter(filterSheet, table.range)
  setFilterCriteria(filterSheet, 3, { kind: 'condition', operator: 'equals', value: 'North' })
  table.autoFilter = filterSheet.autoFilter
  patchSlicerCacheOptions(workbook, slicer, options)
  return { workbook, slicer }
}
