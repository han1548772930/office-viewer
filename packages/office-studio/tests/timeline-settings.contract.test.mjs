import assert from 'node:assert/strict'
import { test } from 'vitest'
import { workbookFromInput } from '../src/api/model-utils.ts'
import { normalizeTimelineSettingsPatch, patchTimelineSettings, readTimelineSettings } from '../src/features/slicer/timeline-settings.ts'
import { timelineCache, timelineFilterWorkbook, timelineViews } from './fixtures/timeline-filter-fixture.mjs'

// Per-view settings and model integration drafts only; every test is unexecuted.
const clone = value => structuredClone(value)
const setup = () => { const workbook = timelineFilterWorkbook(); return { workbook, view: timelineViews(workbook)[0], shared: timelineViews(workbook)[1] } }
const optional = ['caption', 'style', 'scrollPosition', 'showHeader', 'showSelectionLabel', 'showTimeLevel', 'showHorizontalScrollbar']
const switches = ['showHeader', 'showSelectionLabel', 'showTimeLevel', 'showHorizontalScrollbar']

for (const level of [0, 1, 2, 3]) for (const selectionLevel of [0, 1, 2, 3]) test(`public grains ${level}/${selectionLevel} only edit the current view`, () => {
  const f = setup(), before = clone(f.workbook), settings = readTimelineSettings(f.view), changed = settings.level !== level || settings.selectionLevel !== selectionLevel
  assert.equal(patchTimelineSettings(f.workbook, f.view, { level, selectionLevel }), changed)
  assert.equal(readTimelineSettings(f.view).level, level); assert.equal(readTimelineSettings(f.view).selectionLevel, selectionLevel)
  assert.equal(f.view.level, ['year', 'quarter', 'month', 'day'][level]); assert.equal(f.view.selectionLevel, ['year', 'quarter', 'month', 'day'][selectionLevel])
  assert.equal(f.view.scrollPosition, before.slicers[0].scrollPosition)
  assert.equal(f.view.revision, before.slicers[0].revision + Number(changed))
  assert.deepEqual(f.shared, before.slicers[1]); assert.deepEqual(f.workbook.slicerCaches, before.slicerCaches)
  assert.deepEqual(f.workbook.pivotTables, before.pivotTables); assert.deepEqual(f.workbook.charts, before.charts)
})

for (const key of switches) test(`${key} authors a current-view boolean without changing date filtering or independent settings`, () => {
  const f = setup(), before = clone(f.workbook)
  assert.equal(patchTimelineSettings(f.workbook, f.view, { [key]: false }), true)
  assert.equal(readTimelineSettings(f.view)[key], false)
  assert.deepEqual(f.shared, before.slicers[1]); assert.deepEqual(f.workbook.slicerCaches, before.slicerCaches)
  for (const other of switches.filter(other => other !== key)) assert.equal(f.view[other], before.slicers[0][other])
})

for (const scrollPosition of [-1.75, 0, .5, 60.25, 44927.5]) test(`serial navigation ${scrollPosition} retains finite negative and fractional positions`, () => {
  const f = setup(), before = clone(f.workbook)
  assert.equal(patchTimelineSettings(f.workbook, f.view, { scrollPosition }), true)
  assert.equal(readTimelineSettings(f.view).scrollPosition, scrollPosition)
  assert.equal(f.view.level, before.slicers[0].level); assert.equal(f.view.selectionLevel, before.slicers[0].selectionLevel)
  assert.deepEqual(f.shared, before.slicers[1]); assert.deepEqual(f.workbook.slicerCaches, before.slicerCaches)
})

test('one presence-aware patch changes its authored view once and leaves name, caption and style as independent properties', () => {
  const f = setup(), cache = timelineCache(f.workbook)
  cache.timelineDateRange = { startDate: 60.25, endDate: 61, useWholeDay: false }
  const before = clone(f.workbook), patch = { name: 'Authored Timeline', caption: 'Date selection', style: 'Calendar bands',
    level: 2, selectionLevel: 1, scrollPosition: -10.5, showHeader: false, showSelectionLabel: false, showTimeLevel: false, showHorizontalScrollbar: false }
  assert.equal(patchTimelineSettings(f.workbook, f.view, patch), true); assert.deepEqual(readTimelineSettings(f.view), patch)
  assert.equal(f.view.revision, before.slicers[0].revision + 1)
  assert.deepEqual(f.shared, before.slicers[1]); assert.deepEqual(f.workbook.slicerCaches, before.slicerCaches)
  const after = clone(f.workbook); assert.equal(patchTimelineSettings(f.workbook, f.view, patch), false); assert.deepEqual(f.workbook, after)
  assert.equal(patchTimelineSettings(f.workbook, f.view, {}), false); assert.deepEqual(f.workbook, after)
})

for (const key of optional) test(`own ${key}:undefined clears only that optional override while absence preserves it`, () => {
  const f = setup(), authored = key === 'caption' ? 'Authored caption' : key === 'style' ? 'Calendar bands' : key === 'scrollPosition' ? -.5 : false
  patchTimelineSettings(f.workbook, f.view, { [key]: authored }); const before = clone(f.workbook)
  assert.equal(patchTimelineSettings(f.workbook, f.view, {}), false); assert.deepEqual(f.workbook, before)
  assert.equal(patchTimelineSettings(f.workbook, f.view, { [key]: undefined }), true)
  assert.equal(Object.hasOwn(readTimelineSettings(f.view), key), false); assert.equal(Object.hasOwn(f.view, key), false)
  assert.deepEqual(f.shared, before.slicers[1]); assert.deepEqual(f.workbook.slicerCaches, before.slicerCaches)
  const after = clone(f.workbook); assert.equal(patchTimelineSettings(f.workbook, f.view, { [key]: undefined }), false); assert.deepEqual(f.workbook, after)
})

for (const key of ['name', 'level', 'selectionLevel']) test(`required ${key} cannot be cleared or partially commit another setting`, () => {
  const f = setup(), before = clone(f.workbook)
  assert.throws(() => patchTimelineSettings(f.workbook, f.view, { caption: 'Valid caption', [key]: undefined }), error => error.code === 'invalid-target')
  assert.deepEqual(f.workbook, before)
})

for (const patch of [null, [], { name: '' }, { caption: ' ' }, { caption: 'x'.repeat(32768) }, { style: '' }, { style: 'x'.repeat(256) },
  { level: -1 }, { level: 4 }, { level: .5 }, { selectionLevel: 'Month' }, { scrollPosition: NaN }, { scrollPosition: Infinity },
  { scrollPosition: '44927' }, { showHeader: 1 }, { showSelectionLabel: 'yes' }, { showTimeLevel: null }, { showHorizontalScrollbar: {} },
  { cacheId: 'other' }, { timelineDateRange: { startDate: 1, endDate: 2 } }]) {
  test(`invalid settings ${JSON.stringify(patch)?.slice(0, 80)} preserve the entire workbook`, () => {
    const f = setup(), before = clone(f.workbook)
    assert.throws(() => normalizeTimelineSettingsPatch(patch), error => error.code === 'invalid-target')
    assert.throws(() => patchTimelineSettings(f.workbook, f.view, patch), error => error.code === 'invalid-target')
    assert.deepEqual(f.workbook, before)
  })
}

for (const otherKind of ['timeline', 'slicer']) test(`view names remain unique against existing ${otherKind} controls`, () => {
  const f = setup(), other = otherKind === 'timeline' ? f.shared : f.workbook.slicers.find(view => view.type === 'slicer'), before = clone(f.workbook)
  assert.throws(() => patchTimelineSettings(f.workbook, f.view, { name: other.name.toUpperCase() }), error => error.code === 'invalid-target')
  assert.deepEqual(f.workbook, before)
})

test('builtin and workbook custom Timeline styles are available; missing names are refused atomically and clear restores the default', () => {
  const f = setup()
  assert.equal(patchTimelineSettings(f.workbook, f.view, { style: 'TimelineStyleLight1' }), true)
  assert.equal(patchTimelineSettings(f.workbook, f.view, { style: 'Calendar bands' }), true)
  const before = clone(f.workbook)
  assert.throws(() => patchTimelineSettings(f.workbook, f.view, { caption: 'Valid', style: 'SlicerStyleLight1' }), error => error.code === 'invalid-target')
  assert.deepEqual(f.workbook, before)
  assert.equal(patchTimelineSettings(f.workbook, f.view, { style: undefined }), true)
  assert.equal(readTimelineSettings(f.view).style, undefined)
})

test('missing source does not prevent independent retained view settings and returned settings are defensive', () => {
  const f = setup(); f.workbook.sheets = f.workbook.sheets.filter(sheet => sheet.id !== 'data')
  const before = clone(f.workbook)
  assert.equal(patchTimelineSettings(f.workbook, f.view, { caption: 'Retained source', scrollPosition: -.75, level: 3 }), true)
  const settings = readTimelineSettings(f.view); settings.caption = 'External mutation'; settings.level = 0
  assert.equal(f.view.caption, 'Retained source'); assert.equal(f.view.level, 'day')
  assert.deepEqual(f.shared, before.slicers[1]); assert.deepEqual(f.workbook.slicerCaches, before.slicerCaches)
})

test('ordinary Slicer views reject Timeline settings through both reads and patches', () => {
  const f = setup(), slicer = f.workbook.slicers.find(view => view.type === 'slicer'), before = clone(f.workbook)
  assert.throws(() => readTimelineSettings(slicer), error => error.code === 'invalid-target')
  assert.throws(() => patchTimelineSettings(f.workbook, slicer, { level: 0 }), error => error.code === 'invalid-target')
  assert.deepEqual(f.workbook, before)
})

test('JSON reopening preserves retained string grains and public numeric settings without inventing cleared optional values', () => {
  const f = setup(); patchTimelineSettings(f.workbook, f.view, { level: 3, selectionLevel: 0, scrollPosition: -.25, showTimeLevel: false, style: undefined })
  const settings = readTimelineSettings(f.view), reopened = workbookFromInput(JSON.parse(JSON.stringify(f.workbook))), restored = reopened.slicers.find(view => view.id === f.view.id)
  assert.deepEqual(readTimelineSettings(restored), settings)
  assert.equal(restored.level, 'day'); assert.equal(restored.selectionLevel, 'year'); assert.equal(Object.hasOwn(readTimelineSettings(restored), 'style'), false)
})
