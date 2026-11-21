import { expect, test, type Page, type TestInfo } from '@playwright/test'
import { createHash } from 'node:crypto'
import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { inflateRawSync } from 'node:zlib'

type Resource = { path: string; status: number; contentType: string }

function trackResources(page: Page): Resource[] {
  const resources: Resource[] = []
  page.on('response', (response) => {
    const path = new URL(response.url()).pathname
    if (/^\/(?:wasm|wasm-writer|compute)\//u.test(path)) resources.push({ path, status: response.status(), contentType: response.headers()['content-type'] ?? '' })
  })
  return resources
}

/** Inspect actual ZIP entries using Node's independent DEFLATE implementation. */
function officePart(bytes: Buffer, expected: string): string {
  const minimum = Math.max(0, bytes.length - 65_557)
  let end = bytes.length - 22
  while (end >= minimum && bytes.readUInt32LE(end) !== 0x06054b50) end--
  if (end < minimum) throw new Error('Office download has no ZIP central directory')
  let cursor = bytes.readUInt32LE(end + 16)
  const entries = bytes.readUInt16LE(end + 10)
  for (let index = 0; index < entries; index++) {
    if (bytes.readUInt32LE(cursor) !== 0x02014b50) throw new Error('Invalid ZIP central directory entry')
    const method = bytes.readUInt16LE(cursor + 10)
    const size = bytes.readUInt32LE(cursor + 20)
    const nameLength = bytes.readUInt16LE(cursor + 28)
    const extraLength = bytes.readUInt16LE(cursor + 30)
    const commentLength = bytes.readUInt16LE(cursor + 32)
    const name = bytes.subarray(cursor + 46, cursor + 46 + nameLength).toString('utf8')
    if (name === expected) {
      const local = bytes.readUInt32LE(cursor + 42)
      if (bytes.readUInt32LE(local) !== 0x04034b50) throw new Error('Invalid ZIP local entry')
      const offset = local + 30 + bytes.readUInt16LE(local + 26) + bytes.readUInt16LE(local + 28)
      const data = bytes.subarray(offset, offset + size)
      if (method === 0) return data.toString('utf8')
      if (method === 8) return inflateRawSync(data).toString('utf8')
      throw new Error(`Unsupported ZIP method ${method}`)
    }
    cursor += 46 + nameLength + extraLength + commentLength
  }
  throw new Error(`Missing Office part ${expected}`)
}

async function copyDocumentText(page: Page): Promise<string> {
  const bridge = page.getByRole('textbox', { name: '文档文字输入', exact: true })
  await bridge.focus()
  await bridge.press('Control+A')
  const text = await bridge.evaluate((element) => {
    const clipboard = new DataTransfer()
    element.dispatchEvent(new ClipboardEvent('copy', { bubbles: true, cancelable: true, clipboardData: clipboard }))
    return clipboard.getData('text/plain')
  })
  await bridge.press('ArrowLeft')
  return text
}

async function diagnostics(page: Page): Promise<string[]> {
  await page.getByRole('button', { name: '文件菜单', exact: true }).click()
  await page.getByRole('menuitem', { name: /^解析诊断/u }).click()
  const dialog = page.getByRole('dialog', { name: '解析诊断', exact: true })
  await expect(dialog).toBeVisible()
  const messages = await dialog.locator('.workspace-diagnostics li').allTextContents()
  await dialog.locator('.ant-modal-close').click()
  await expect(dialog).toBeHidden()
  return messages
}

async function saveDownload(page: Page): Promise<{ bytes: Buffer; name: string }> {
  const downloadPromise = page.waitForEvent('download')
  await page.getByRole('button', { name: '保存 Office 文件', exact: true }).click()
  const download = await downloadPromise
  const path = await download.path()
  if (!path) throw new Error('The Office download has no local artifact')
  const bytes = await readFile(path)
  expect(bytes.subarray(0, 4).toString('hex')).toBe('504b0304')
  return { bytes, name: download.suggestedFilename() }
}

async function attachEvidence(testInfo: TestInfo, value: Record<string, unknown>): Promise<void> {
  await testInfo.attach('production-office-evidence.json', { body: Buffer.from(JSON.stringify(value, null, 2)), contentType: 'application/json' })
}

function expectRealParserAndWriter(resources: Resource[]): void {
  for (const path of ['/wasm/office_parser.js', '/wasm/office_parser_bg.wasm', '/wasm-writer/office_writer.js', '/wasm-writer/office_writer_bg.wasm']) {
    expect(resources.some((resource) => resource.path === path && resource.status === 200), path).toBe(true)
  }
}

test('production DOCX import, Canvas edit, real file download and reimport retain edited text', async ({ page }, testInfo) => {
  const resources = trackResources(page)
  const errors: string[] = []
  page.on('pageerror', (error) => errors.push(error.message))
  const sourcePath = fileURLToPath(new URL('../fixtures/exchange/exchange-synthetic-docs.docx', import.meta.url))
  const source = await readFile(sourcePath)
  await page.goto('/')
  await page.getByLabel('选择 Office 文件', { exact: true }).setInputFiles(sourcePath)
  await expect(page.locator('.document-workbench')).toBeVisible()
  await expect(page.getByRole('button', { name: '重命名当前文件', exact: true })).toHaveText('exchange-synthetic-docs.docx')
  await expect.poll(() => copyDocumentText(page)).toContain('Exchange Synthetic Document')
  const before = await copyDocumentText(page)
  expect(before).toContain('Fictional data table')
  const importDiagnostics = await diagnostics(page)
  const marker = '生产文档回环验证-'
  const bridge = page.getByRole('textbox', { name: '文档文字输入', exact: true })
  await bridge.focus()
  await bridge.press('Control+Home')
  await bridge.pressSequentially(marker)
  await expect.poll(() => copyDocumentText(page)).toContain(marker)
  const downloaded = await saveDownload(page)
  expect(downloaded.name).toMatch(/\.docx$/iu)
  const body = officePart(downloaded.bytes, 'word/document.xml')
  expect(body.replace(/<[^>]*>/gu, '')).toContain(marker)
  expect(body).toContain('Exchange Synthetic Document')
  expect(officePart(downloaded.bytes, '[Content_Types].xml')).toContain('wordprocessingml.document.main+xml')
  await page.getByLabel('选择 Office 文件', { exact: true }).setInputFiles({ name: downloaded.name, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: downloaded.bytes })
  await expect(page.locator('.workspace-task-strip')).toBeHidden()
  await expect.poll(() => copyDocumentText(page)).toContain(marker)
  const reimported = await copyDocumentText(page)
  expect(reimported).toContain('Fictional data table')
  expectRealParserAndWriter(resources)
  expect(errors).toEqual([])
  await attachEvidence(testInfo, { format: 'DOCX', sourceSha256: createHash('sha256').update(source).digest('hex'), downloadSha256: createHash('sha256').update(downloaded.bytes).digest('hex'), downloadedBytes: downloaded.bytes.length, importDiagnostics, reimportDiagnostics: await diagnostics(page), resources, before, reimported, independentReader: 'Node ZIP/DEFLATE XML inspection; full Office application fidelity remains unverified' })
})

test('production DOCX numbering roundtrip retains unchecked and checked checklist labels', async ({ page }) => {
  await page.goto('/')
  const labels = await page.evaluate(async () => {
    const writerURL = new URL('/wasm-writer/office_writer.js', location.href).href
    const parserURL = new URL('/wasm/office_parser.js', location.href).href
    const writer = await import(writerURL)
    const parser = await import(parserURL)
    await writer.default()
    await parser.default()
    const model = {
      kind: 'document',
      name: 'Checklist roundtrip',
      blocks: [
        { id: 'unchecked', type: 'checklist', level: 0, checklist: { listId: 'tasks', checked: false }, runs: [{ text: 'Prepare files' }] },
        { id: 'checked', type: 'checklist', level: 0, checklist: { listId: 'tasks', checked: true }, runs: [{ text: 'Review changes' }] },
      ],
    }
    const output = writer.write_office('document', JSON.stringify(model))
    try {
      const parsed = JSON.parse(parser.parse_document(output.bytes, 'checklist-roundtrip.docx'))
      return parsed.payload.data.blocks.map((block: { details?: { listLabel?: string } }) => block.details?.listLabel)
    } finally { output.free() }
  })
  expect(labels).toEqual(['☐', '☑'])
})

test('production XLSX import, pending formula-bar edit, save and reimport retain cell data and formulas', async ({ page }, testInfo) => {
  const resources = trackResources(page)
  const errors: string[] = []
  page.on('pageerror', (error) => errors.push(error.message))
  const sourcePath = fileURLToPath(new URL('../fixtures/exchange/exchange-synthetic-sheets.xlsx', import.meta.url))
  const source = await readFile(sourcePath)
  await page.goto('/')
  await page.getByLabel('选择 Office 文件', { exact: true }).setInputFiles(sourcePath)
  await expect(page.getByRole('button', { name: '重命名当前文件', exact: true })).toHaveText('exchange-synthetic-sheets.xlsx')
  await expect(page.getByRole('tab', { name: 'Synthetic Data', exact: true })).toHaveAttribute('aria-selected', 'true')
  const address = page.getByRole('textbox', { name: '名称框', exact: true })
  const formula = page.getByRole('textbox', { name: '公式编辑', exact: true })
  await address.fill('A2')
  await address.press('Enter')
  await expect(formula).toHaveValue('Amber Widget')
  const importDiagnostics = await diagnostics(page)
  const marker = '生产表格回环验证'
  await formula.fill(marker)
  // Save must commit the live formula-bar session without a prior Enter or blur.
  const downloaded = await saveDownload(page)
  expect(downloaded.name).toMatch(/\.xlsx$/iu)
  const sheet = officePart(downloaded.bytes, 'xl/worksheets/sheet1.xml')
  const strings = (() => { try { return officePart(downloaded.bytes, 'xl/sharedStrings.xml') } catch { return '' } })()
  expect(`${sheet}${strings}`).toContain(marker)
  expect(sheet).toContain('B2*C2')
  expect(officePart(downloaded.bytes, 'xl/workbook.xml')).toContain('Synthetic Data')
  await page.getByLabel('选择 Office 文件', { exact: true }).setInputFiles({ name: downloaded.name, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: downloaded.bytes })
  await expect(page.locator('.workspace-task-strip')).toBeHidden()
  await address.fill('A2')
  await address.press('Enter')
  await expect(formula).toHaveValue(marker)
  await address.fill('D2')
  await address.press('Enter')
  await expect(formula).toHaveValue('=B2*C2')
  expectRealParserAndWriter(resources)
  expect(errors).toEqual([])
  await attachEvidence(testInfo, { format: 'XLSX', sourceSha256: createHash('sha256').update(source).digest('hex'), downloadSha256: createHash('sha256').update(downloaded.bytes).digest('hex'), downloadedBytes: downloaded.bytes.length, importDiagnostics, reimportDiagnostics: await diagnostics(page), resources, verifiedCells: { A2: marker, D2: '=B2*C2' }, independentReader: 'Node ZIP/DEFLATE XML inspection; full Office application fidelity remains unverified' })
})

test('production computation assets initialize real Rust WASM in a browser worker', async ({ page }, testInfo) => {
  const resources = trackResources(page)
  await page.goto('/')
  const result = await page.evaluate(async () => {
    const code = 'self.onmessage = async ({ data }) => { let module; try { module = await import(data); await module.default(); const aggregate = new module.NumericAggregator(); try { aggregate.push(new Float64Array([4, 7, 9]), new Uint8Array([1, 1, 1])); self.postMessage({ version: module.compute_version(), values: Array.from(aggregate.result()) }); } finally { aggregate.free(); } } catch(error) { self.postMessage({ error: String(error) }); } }'
    const url = URL.createObjectURL(new Blob([code], { type: 'text/javascript' }))
    const worker = new Worker(url, { type: 'module' })
    try {
      return await new Promise<{ version?: string; values?: number[]; error?: string }>((resolve, reject) => {
        const timeout = window.setTimeout(() => reject(new Error('Production compute worker timed out')), 20_000)
        worker.onmessage = ({ data }) => { window.clearTimeout(timeout); resolve(data) }
        worker.onerror = (event) => { window.clearTimeout(timeout); reject(new Error(event.message)) }
        worker.postMessage(new URL('/compute/office_compute.js', location.origin).href)
      })
    } finally { worker.terminate(); URL.revokeObjectURL(url) }
  })
  expect(result.error).toBeUndefined()
  expect(result.version).toMatch(/^\d+\.\d+\.\d+/u)
  expect(result.values?.[0]).toBe(3)
  expect(result.values?.[5]).toBe(20)
  for (const path of ['/compute/office_compute.js', '/compute/office_compute_bg.wasm']) expect(resources.some((resource) => resource.path === path && resource.status === 200), path).toBe(true)
  await attachEvidence(testInfo, { runtime: 'production preview / dedicated Worker / real public Rust WASM', result, resources, limitation: 'Confirms artifact serving and execution; does not measure large-file performance or prove every product computation uses this module.' })
})
