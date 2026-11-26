import { expect, test } from '@playwright/test'
import { inflateRawSync } from 'node:zlib'

// Real Worker/WASM/Facade/browser drafts. Requires rebuilt parser/writer
// artifacts in the later unified pass; no browser or server was started here.
function worksheet(values: number[]): string {
  const bytes = Buffer.from(values)
  let end = bytes.length - 22
  while (end >= Math.max(0, bytes.length - 65_557) && bytes.readUInt32LE(end) !== 0x06054b50) end--
  if (end < 0) throw new Error('XLSX has no central directory')
  let cursor = bytes.readUInt32LE(end + 16)
  const count = bytes.readUInt16LE(end + 10)
  for (let index = 0; index < count; index++) {
    if (bytes.readUInt32LE(cursor) !== 0x02014b50) throw new Error('Invalid XLSX central directory')
    const method = bytes.readUInt16LE(cursor + 10), size = bytes.readUInt32LE(cursor + 20)
    const nameLength = bytes.readUInt16LE(cursor + 28), extraLength = bytes.readUInt16LE(cursor + 30), commentLength = bytes.readUInt16LE(cursor + 32)
    const name = bytes.subarray(cursor + 46, cursor + 46 + nameLength).toString('utf8')
    if (name === 'xl/worksheets/sheet1.xml') {
      const local = bytes.readUInt32LE(cursor + 42)
      if (bytes.readUInt32LE(local) !== 0x04034b50) throw new Error('Invalid XLSX local directory')
      const offset = local + 30 + bytes.readUInt16LE(local + 26) + bytes.readUInt16LE(local + 28)
      const data = bytes.subarray(offset, offset + size)
      if (method === 0) return data.toString('utf8')
      if (method === 8) return inflateRawSync(data).toString('utf8')
      throw new Error('Unsupported XLSX compression')
    }
    cursor += 46 + nameLength + extraLength + commentLength
  }
  throw new Error('XLSX has no worksheet')
}
async function roundtrip(page: import('@playwright/test').Page, options: { boundaryEight?: boolean } = {}) {
  await page.goto('/packages/office-studio/tests/fixtures/sheet-outline-office-io.html')
  await page.waitForFunction(() => 'outlineOfficeIo' in window)
  return page.evaluate(async options => (window as unknown as { outlineOfficeIo: { roundtrip(options: unknown): Promise<any> } }).outlineOfficeIo.roundtrip(options), options)
}

test('actual Facade edits survive Worker/WASM XLSX write, parser normalization and reopening', async ({ page }, testInfo) => {
  const value = await roundtrip(page), xml = worksheet(value.bytes)
  expect(value.actualGroups).toEqual(value.expectedGroups)
  expect(value.first.sheets[0].outline).toEqual(value.before.sheets[0].outline)
  expect(value.first.sheets[0].rowCount).toBe(32); expect(value.first.sheets[0].columnCount).toBe(12)
  expect(xml.match(/<sheetPr\b/gu)).toHaveLength(1)
  expect(xml).toContain('outlineLevelRow="2"'); expect(xml).toContain('outlineLevelCol="1"')
  expect(xml).toContain('summaryBelow="1"'); expect(xml).toContain('summaryRight="0"')
  expect(xml).toContain('urn:office-studio:sheet-outline:1')
  expect(value.write.metrics.boundary.outputCopies).toBe(1)
  expect(value.parse.task.metrics.transport.inputTransfers).toBe(1)
  await testInfo.attach('outline-actual-io.json', { body: Buffer.from(JSON.stringify({ write: value.write, parse: value.parse, firstOutline: value.first.sheets[0].outline }, null, 2)), contentType: 'application/json' })
})

test('reopened Facade undo redo and another binary save keep manual overlap and public plugin identity', async ({ page }) => {
  const value = await roundtrip(page)
  expect(value.undoneGroups).toEqual(value.actualGroups); expect(value.redoneGroups).toEqual(value.expandedGroups)
  expect(value.expandedGroups.find((group: { axis: string }) => group.axis === 'row').collapsed).toBe(false)
  for (const key of ['hiddenRows', 'hiddenColumns', 'rowHiddenSpans', 'columnHiddenSpans']) {
    expect(value.first.sheets[0][key]).toEqual(value.before.sheets[0][key])
    expect(value.after.sheets[0][key]).toEqual(value.before.sheets[0][key])
  }
  expect(value.after.sheets[0].outline).toEqual(value.expanded.sheets[0].outline)
  expect(value.after.sheets[0].outline.groups[0].futureGroup).toEqual({ text: 'A & B' })
  const plugins = value.publicSnapshot.resources.filter((resource: { name: string }) => resource.name === 'SHEET_OUTLINE_PLUGIN')
  expect(plugins).toHaveLength(1)
  const resource = typeof plugins[0].data === 'string' ? JSON.parse(plugins[0].data) : plugins[0].data
  expect(Object.values(resource).flatMap((sheet: any) => sheet.outlineList).map((group: { id: string }) => group.id).sort())
    .toEqual(value.expectedGroups.map((group: { id: string }) => group.id).sort())
})

test('eighth model layer and the full-axis boundary reopen intact with explicit native projection diagnostics', async ({ page }) => {
  const value = await roundtrip(page, { boundaryEight: true }), xml = worksheet(value.bytes)
  expect(value.first.sheets[0].outline.groups.filter((group: { axis: string }) => group.axis === 'row')).toHaveLength(8)
  expect(value.first.sheets[0].outline).toEqual(value.before.sheets[0].outline)
  expect(value.first.sheets[0].rowCount).toBe(32)
  expect(xml).toContain('outlineLevelRow="7"')
  expect(value.write.warnings.some((warning: { message: string }) => warning.message.includes('OUTLINE_DEPTH_SNAPSHOT'))).toBe(true)
  expect(value.write.warnings.some((warning: { message: string }) => warning.message.includes('OUTLINE_BOUNDARY_SNAPSHOT'))).toBe(true)
})
