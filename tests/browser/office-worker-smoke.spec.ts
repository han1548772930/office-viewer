import { expect, test } from '@playwright/test'
import { inflateRawSync } from 'node:zlib'
import { writeFile } from 'node:fs/promises'
import { wasmTaskFixture } from './wasm-task-fixtures.ts'

function officeXml(bytes: Buffer, name: string): string {
  let end = bytes.length - 22
  const minimum = Math.max(0, bytes.length - 65_557)
  while (end >= minimum && bytes.readUInt32LE(end) !== 0x06054b50) end--
  if (end < minimum) throw new Error('The Office file has no ZIP central directory')
  let cursor = bytes.readUInt32LE(end + 16)
  for (let index = 0; index < bytes.readUInt16LE(end + 10); index++) {
    if (bytes.readUInt32LE(cursor) !== 0x02014b50) throw new Error('Invalid ZIP central entry')
    const method = bytes.readUInt16LE(cursor + 10)
    const size = bytes.readUInt32LE(cursor + 20)
    const length = bytes.readUInt16LE(cursor + 28)
    const path = bytes.subarray(cursor + 46, cursor + 46 + length).toString('utf8')
    if (path === name) {
      const local = bytes.readUInt32LE(cursor + 42)
      const offset = local + 30 + bytes.readUInt16LE(local + 26) + bytes.readUInt16LE(local + 28)
      const content = bytes.subarray(offset, offset + size)
      if (method === 0) return content.toString('utf8')
      if (method === 8) return inflateRawSync(content).toString('utf8')
      throw new Error(`Unsupported ZIP compression ${method}`)
    }
    cursor += 46 + length + bytes.readUInt16LE(cursor + 30) + bytes.readUInt16LE(cursor + 32)
  }
  throw new Error(`Missing Office part ${name}`)
}

/** This uses the shipped JS/WASM inside real module Workers, without a mock or editor UI. */
test('production Office WASM writes DOCX and XLSX without requiring a host clock', async ({ page }) => {
  await page.goto('/')
  const results = await page.evaluate(async (context) => {
    const moduleURL = new URL('/wasm-writer/office_writer.js', location.href).href
    const source = `
      try {
        const writer = await import(${JSON.stringify(moduleURL)});
        const initialized = await writer.default();
        if (writer.writer_protocol_version() !== 2) throw new Error('Writer WASM protocol mismatch');
        const context = ${JSON.stringify(context)};
        const outputs = [];
        for (const [kind, model] of [
          ['document', { kind: 'document', name: 'Worker DOCX', blocks: [{ type: 'paragraph', runs: [{ text: 'Browser WASM DOCX' }] }] }],
          ['workbook', { kind: 'workbook', name: 'Worker XLSX', sheets: [{ id: 'data', name: 'Data', cells: { '0:0': { row: 0, col: 0, value: 'Browser WASM XLSX' } } }] }],
        ]) {
          const first = await writer.write_office(kind, JSON.stringify(model), { ...context, requestId: kind + '-first' });
          let second;
          try {
            second = await writer.write_office(kind, JSON.stringify(model), { ...context, requestId: kind + '-second' });
            const a = first.copy_bytes();
            const b = second.copy_bytes();
            const metadata = first.metadata();
            if (a.buffer === initialized.memory.buffer || a.byteOffset !== 0 || a.byteLength !== a.buffer.byteLength) throw new Error('Writer returned aliased bytes');
            outputs.push({ kind, size: a.byteLength, signature: Array.from(a.slice(0, 4)), deterministic: a.length === b.length && a.every((value, index) => value === b[index]), warnings: metadata.warnings });
          } finally { first.free(); second?.free(); }
        }
        postMessage({ ok: true, outputs });
      } catch (error) { postMessage({ ok: false, error: String(error) }); }
    `
    const url = URL.createObjectURL(new Blob([source], { type: 'application/javascript' }))
    const worker = new Worker(url, { type: 'module' })
    try {
      return await new Promise<{ ok: boolean; error?: string; outputs?: Array<{ kind: string; size: number; signature: number[]; deterministic: boolean; warnings: unknown[] }> }>((resolve, reject) => {
        const timeout = setTimeout(() => reject(new Error('The Office writer Worker did not respond')), 15_000)
        worker.onmessage = ({ data }) => { clearTimeout(timeout); resolve(data) }
        worker.onerror = (error) => { clearTimeout(timeout); reject(new Error(error.message)) }
      })
    } finally { worker.terminate(); URL.revokeObjectURL(url) }
  }, wasmTaskFixture('export', 'worker-deterministic'))
  expect(results.ok, results.error).toBe(true)
  expect(results.outputs?.map((output) => output.kind)).toEqual(['document', 'workbook'])
  for (const output of results.outputs ?? []) {
    expect(output.size).toBeGreaterThan(500)
    expect(output.signature).toEqual([0x50, 0x4b, 0x03, 0x04])
    expect(output.deterministic).toBe(true)
    expect(output.warnings).toEqual([])
  }
})

test('production XLSX Worker preserves AutoFilter criteria and both filter/manual hidden rows', async ({ page }, testInfo) => {
  await page.goto('/')
  const result = await page.evaluate(async (context) => {
    const moduleURL = new URL('/wasm-writer/office_writer.js', location.href).href
    const model = { kind: 'workbook', name: 'AutoFilter', sheets: [{ id: 'filtered', name: 'Data', cells: {
      B3: { value: 'Region' }, D3: { value: 'Revenue' }, B4: { value: 'East' }, D4: { value: 150 },
      B5: { value: 'West' }, D5: { value: 250 }, B6: { value: 'East' }, D6: { value: 100 }, B7: { value: 'West' }, D7: { value: 50 },
    }, hiddenRows: [5], autoFilter: {
      range: { startRow: 2, startCol: 1, endRow: 6, endCol: 3 },
      columns: { 1: { kind: 'values', selectedItems: ['East'] }, 3: { kind: 'condition', operator: 'between', value: 100, secondValue: 200 } },
      filteredRows: [4, 6],
    } }] }
    const script = `
      onmessage = async ({ data }) => {
        try {
          const writer = await import(${JSON.stringify(moduleURL)});
          const initialized = await writer.default();
          if (writer.writer_protocol_version() !== 2) throw new Error('Writer WASM protocol mismatch');
          const result = await writer.write_office('workbook', JSON.stringify(data), ${JSON.stringify(context)});
          try {
            const bytes = result.copy_bytes();
            const metadata = result.metadata();
            if (bytes.buffer === initialized.memory.buffer || bytes.byteOffset !== 0 || bytes.byteLength !== bytes.buffer.byteLength) throw new Error('Writer returned aliased bytes');
            postMessage({ bytes, warnings: metadata.warnings }, [bytes.buffer]);
          }
          finally { result.free(); }
        } catch (error) { postMessage({ error: String(error) }); }
      };
    `
    const url = URL.createObjectURL(new Blob([script], { type: 'application/javascript' }))
    const worker = new Worker(url, { type: 'module' })
    try {
      return await new Promise<{ bytes?: number[]; warnings?: unknown[]; error?: string }>((resolve, reject) => {
        const timeout = setTimeout(() => reject(new Error('The filtered XLSX Writer did not respond')), 15_000)
        worker.onmessage = ({ data }) => { clearTimeout(timeout); resolve({ ...data, bytes: data.bytes ? Array.from(data.bytes) : undefined }) }
        worker.onerror = (error) => { clearTimeout(timeout); reject(new Error(error.message)) }
        worker.postMessage(model)
      })
    } finally { worker.terminate(); URL.revokeObjectURL(url) }
  }, wasmTaskFixture('export', 'worker-auto-filter'))
  expect(result.error).toBeUndefined()
  expect(result.warnings).toEqual([])
  const bytes = Buffer.from(result.bytes!)
  const worksheet = officeXml(bytes, 'xl/worksheets/sheet1.xml')
  expect(worksheet).toContain('<sheetPr filterMode="1"/>')
  expect(worksheet).toContain('<autoFilter ref="B3:D7"><filterColumn colId="0"><filters blank="0"><filter val="East"/></filters></filterColumn>')
  expect(worksheet).toContain('<filterColumn colId="2"><customFilters and="1"><customFilter operator="greaterThanOrEqual" val="100"/><customFilter operator="lessThanOrEqual" val="200"/></customFilters></filterColumn>')
  for (const row of [5, 6, 7]) expect(worksheet).toContain(`<row r="${row}" hidden="1">`)
  expect(worksheet).not.toContain('<row r="4" hidden="1">')
  const artifact = testInfo.outputPath('auto-filter-export.xlsx')
  await writeFile(artifact, bytes)
  await testInfo.attach('auto-filter-export.xlsx', { path: artifact, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
})
