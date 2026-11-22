import { expect, test, type Page } from '@playwright/test'
import { createHash } from 'node:crypto'
import { readFile } from 'node:fs/promises'
import { inflateRawSync } from 'node:zlib'

interface ParsedDocument {
  blocks: Array<{ content: string; details: { inlines: Array<{ text?: string; style?: { fontFamilyAscii?: string; fontFamilyEastAsia?: string; fontSizePt?: number; bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean; color?: string; highlight?: string; backgroundColor?: string; verticalAlign?: string } }>; officeNumberingMarker?: { text: string } } }>
  comments: Array<{ text: string; targets: Array<{ start: number; end: number }> }>
  officeSource?: { version: number; styles?: { partName: string; xml: string } }
  defaultStyle?: { version: number; text?: { fontFamilyAscii?: string; fontFamilyEastAsia?: string; fontSizePt?: number; highlight?: string; backgroundColor?: string }; paragraph?: { align?: string; spacing?: { lineRule?: string; lineTwips?: number; afterTwips?: number } } }
}

interface WorkerResult { bytes?: number[]; data?: ParsedDocument; warnings?: Array<{ code: string; message: string }> }

/** Fixture creation and independent reparse use real public WASM in a dedicated Worker. */
async function officeWorker(page: Page, request: { action: 'write'; model: string } | { action: 'parse'; bytes: number[] }): Promise<WorkerResult> {
  return page.evaluate(async (request) => {
    const code = `self.onmessage = async ({ data }) => {
      try {
        const module = await import(data.moduleURL);
        await module.default();
        if (data.action === 'write') {
          const result = module.write_office('document', data.model);
          try { const bytes = result.bytes; self.postMessage({ bytes: bytes.buffer, warnings: JSON.parse(result.warnings) }, [bytes.buffer]); }
          finally { result.free(); }
        } else {
          const envelope = JSON.parse(module.parse_document(new Uint8Array(data.bytes), 'numbering-roundtrip.docx'));
          self.postMessage({ envelope });
        }
      } catch (error) { self.postMessage({ error: String(error) }); }
    }`
    const url = URL.createObjectURL(new Blob([code], { type: 'text/javascript' }))
    const worker = new Worker(url, { type: 'module', name: 'docx-numbering-evidence' })
    let timeout: number | undefined
    try {
      return await new Promise<WorkerResult>((resolve, reject) => {
        timeout = window.setTimeout(() => reject(new Error('DOCX evidence Worker timed out')), 20_000)
        worker.onerror = (event) => reject(new Error(event.message))
        worker.onmessage = ({ data }: MessageEvent<{ bytes?: ArrayBuffer; warnings?: WorkerResult['warnings']; envelope?: { payload: { data: ParsedDocument } }; error?: string }>) => {
          if (data.error) reject(new Error(data.error))
          else resolve({ bytes: data.bytes ? Array.from(new Uint8Array(data.bytes)) : undefined, warnings: data.warnings, data: data.envelope?.payload.data })
        }
        worker.postMessage({ ...request, moduleURL: new URL(request.action === 'write' ? '/wasm-writer/office_writer.js' : '/wasm/office_parser.js', location.origin).href })
      })
    } finally { window.clearTimeout(timeout); worker.terminate(); URL.revokeObjectURL(url) }
  }, request)
}

function sourceModel(): Record<string, unknown> {
  const properties: Record<string, unknown> = Object.fromEntries(['bold', 'italic', 'underline', 'strike', 'fontFamily', 'eastAsiaFontFamily', 'fontSize', 'color', 'verticalAlign', 'href', 'footnoteId', 'endnoteId'].map((name) => [name, null]))
  properties.metadata = Object.fromEntries(['highlight', 'verticalAlign', 'href', 'link', 'note', 'footnoteId', 'endnoteId', 'latexFormula'].map((name) => [name, null]))
  const marker = '<w:bookmarkStart w:id="99" w:name="office_list_v1_browser"/><w:r><w:t xml:space="preserve">[</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ office_seq_browser \\r 0 </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:t xml:space="preserve">0</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve">] </w:t></w:r><w:bookmarkEnd w:id="99"/>'
  const styles = '<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="Original Marker Base"/><w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:rFonts w:ascii="Consolas"/><w:color w:val="124E78"/><w:b/></w:rPr></w:style><w:style w:type="character" w:styleId="EndnoteReference"><w:name w:val="Original Marker"/><w:basedOn w:val="FootnoteReference"/><w:rPr><w:sz w:val="28"/><w:i/></w:rPr></w:style></w:styles>'
  return {
    kind: 'document', name: 'Numbering evidence',
    blocks: [{ id: 'marker', type: 'paragraph', level: 0, runs: [{ id: 'source', text: '[0] Body' }], metadata: { numbering: {}, officeNumberingMarker: {
      version: 1, bookmark: 'office_list_v1_browser', bookmarkId: 99, namespaces: { w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' }, xml: marker, text: '[0] ',
      baseline: { version: 1, numbering: { kind: null, level: null, list: null, checklist: null, definition: null, listId: null, numbering: {} }, prefix: [{ text: '[0] ', properties }] },
    } } }],
    metadata: { source: { officeSource: { version: 1, styles: { partName: 'word/styles.xml', xml: styles } } } },
    comments: { label: { anchor: [{ blockId: 'marker', startOffset: 1, endOffset: 3 }], messages: [{ id: 'message', authorName: 'Fixture Author', text: 'Keep label review', createdAt: '2026-11-22T00:00:00Z' }], resolved: false } },
  }
}

function officePart(bytes: Buffer, path: string): string {
  let end = bytes.length - 22
  const minimum = Math.max(0, bytes.length - 65_557)
  while (end >= minimum && bytes.readUInt32LE(end) !== 0x06054b50) end--
  if (end < minimum) throw new Error('Office download has no ZIP central directory')
  let cursor = bytes.readUInt32LE(end + 16)
  const entries = bytes.readUInt16LE(end + 10)
  for (let index = 0; index < entries; index++) {
    if (bytes.readUInt32LE(cursor) !== 0x02014b50) throw new Error('Invalid ZIP central directory')
    const method = bytes.readUInt16LE(cursor + 10), size = bytes.readUInt32LE(cursor + 20)
    const nameLength = bytes.readUInt16LE(cursor + 28), extraLength = bytes.readUInt16LE(cursor + 30), commentLength = bytes.readUInt16LE(cursor + 32)
    if (bytes.subarray(cursor + 46, cursor + 46 + nameLength).toString('utf8') === path) {
      const local = bytes.readUInt32LE(cursor + 42)
      if (bytes.readUInt32LE(local) !== 0x04034b50) throw new Error('Invalid ZIP local header')
      const offset = local + 30 + bytes.readUInt16LE(local + 26) + bytes.readUInt16LE(local + 28)
      const data = bytes.subarray(offset, offset + size)
      if (method === 0) return data.toString('utf8')
      if (method === 8) return inflateRawSync(data).toString('utf8')
      throw new Error(`Unsupported ZIP compression ${method}`)
    }
    cursor += 46 + nameLength + extraLength + commentLength
  }
  throw new Error(`Missing Office part ${path}`)
}

async function upload(page: Page, bytes: Buffer, name = 'numbering-evidence.docx'): Promise<void> {
  await page.getByLabel('选择 Office 文件', { exact: true }).setInputFiles({ name, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: bytes })
  await expect(page.locator('.document-workbench')).toBeVisible()
  await expect(page.getByRole('button', { name: '重命名当前文件', exact: true })).toHaveText(name)
  await expect(page.locator('.workspace-task-strip')).toBeHidden()
}

async function copyText(page: Page): Promise<string> {
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

async function download(page: Page): Promise<Buffer> {
  const promise = page.waitForEvent('download')
  await page.getByRole('button', { name: '保存 Office 文件', exact: true }).click()
  const result = await promise, path = await result.path()
  if (!path) throw new Error('DOCX download has no artifact')
  const bytes = await readFile(path)
  expect(bytes.subarray(0, 4).toString('hex')).toBe('504b0304')
  return bytes
}

async function parse(page: Page, bytes: Buffer): Promise<ParsedDocument> {
  const result = await officeWorker(page, { action: 'parse', bytes: [...bytes] })
  if (!result.data) throw new Error('DOCX parser did not return its document DTO')
  return result.data
}

async function source(page: Page): Promise<Buffer> {
  const result = await officeWorker(page, { action: 'write', model: JSON.stringify(sourceModel()) })
  expect(result.warnings).toEqual([])
  if (!result.bytes) throw new Error('DOCX fixture writer did not return bytes')
  return Buffer.from(result.bytes)
}

test('production import, Canvas editing and repeated DOCX saves preserve cached character styles and label comments', async ({ page }, testInfo) => {
  const workers: string[] = [], errors: string[] = []
  page.on('worker', (worker) => workers.push(worker.url()))
  page.on('pageerror', (error) => errors.push(error.message))
  await page.goto('/')
  const input = await source(page)
  await upload(page, input)
  await expect.poll(() => copyText(page)).toContain('[0] Body')
  const bridge = page.getByRole('textbox', { name: '文档文字输入', exact: true })
  await bridge.press('Control+End')
  await bridge.pressSequentially(' edited')
  await expect.poll(() => copyText(page)).toContain('[0] Body edited')
  const downloads: Array<{ sha256: string; bytes: number }> = []
  for (let iteration = 0; iteration < 2; iteration++) {
    const saved = await download(page)
    const styles = officePart(saved, 'word/styles.xml'), document = officePart(saved, 'word/document.xml')
    expect(styles).toContain('Original Marker Base')
    expect(styles).toContain('Original Marker')
    expect(styles).toContain('w:styleId="FootnoteReference"')
    expect(styles).toContain('<w:vertAlign w:val="superscript"/>')
    expect(document).toContain('SEQ office_seq_browser')
    expect(document).toContain('Body edited')
    const parsed = await parse(page, saved)
    expect(parsed.blocks[0]?.content).toBe('[0] Body edited')
    expect(parsed.blocks[0]?.details.officeNumberingMarker?.text).toBe('[0] ')
    const numeral = parsed.blocks[0]?.details.inlines.find((inline) => inline.text === '0')
    expect(numeral?.style?.fontFamilyAscii).toBe('Consolas')
    expect(numeral?.style?.fontSizePt).toBe(14)
    expect(numeral?.style?.bold).toBe(true)
    expect(numeral?.style?.italic).toBe(true)
    expect(numeral?.style?.color?.replace(/^#/u, '').toUpperCase()).toBe('124E78')
    expect(parsed.comments.map((comment) => ({ text: comment.text, ranges: comment.targets.map((target) => [target.start, target.end]) }))).toEqual([{ text: 'Keep label review', ranges: [[1, 3]] }])
    expect(parsed.officeSource?.styles?.xml).toContain('Original Marker Base')
    downloads.push({ sha256: createHash('sha256').update(saved).digest('hex'), bytes: saved.length })
    await upload(page, saved, `numbering-evidence-reopened-${iteration + 1}.docx`)
    await expect.poll(() => copyText(page)).toContain('[0] Body edited')
  }
  expect(workers.some((url) => /\/assets\/parser\.worker-/u.test(url))).toBe(true)
  expect(workers.some((url) => /\/assets\/export\.worker-/u.test(url))).toBe(true)
  expect(errors).toEqual([])
  await testInfo.attach('numbering-docx-evidence.json', { body: Buffer.from(JSON.stringify({ sourceSha256: createHash('sha256').update(input).digest('hex'), downloads, workers, limitation: 'Real production Workers and OPC reparse; independent Office applications and large-file performance remain unverified.' }, null, 2)), contentType: 'application/json' })
})

test('Ribbon replacement, history, real DOCX download and reimport retain a deleted label comment as a point', async ({ page }) => {
  await page.goto('/')
  await upload(page, await source(page))
  await expect.poll(() => copyText(page)).toContain('[0] Body')
  await page.getByRole('button', { name: '有序列表选项', exact: true }).click()
  await page.locator('.doc-list-picker:visible [data-list-preset="ORDER_LIST"]').click()
  const bridge = page.getByRole('textbox', { name: '文档文字输入', exact: true })
  await expect(bridge).toBeFocused()
  await expect.poll(() => copyText(page)).toBe('Body')
  await page.locator('[data-ribbon-item="doc-undo"]:visible').getByRole('button').click()
  await expect.poll(() => copyText(page)).toContain('[0] Body')
  await page.locator('[data-ribbon-item="doc-redo"]:visible').getByRole('button').click()
  await expect.poll(() => copyText(page)).toBe('Body')
  const saved = await download(page), document = officePart(saved, 'word/document.xml')
  expect(document).not.toContain('SEQ office_seq_browser')
  expect(document).toContain('<w:numPr>')
  const start = document.indexOf('<w:commentRangeStart'), end = document.indexOf('<w:commentRangeEnd'), reference = document.indexOf('<w:commentReference')
  for (const position of [start, end, reference]) expect(position).toBeGreaterThanOrEqual(0)
  expect(start).toBeLessThan(end)
  expect(end).toBeLessThan(reference)
  const parsed = await parse(page, saved)
  expect(parsed.blocks[0]?.content).toBe('Body')
  expect(parsed.comments.map((comment) => ({ text: comment.text, ranges: comment.targets.map((target) => [target.start, target.end]) }))).toEqual([{ text: 'Keep label review', ranges: [[0, 0]] }])
  await upload(page, saved, 'numbering-point-reopened.docx')
  await expect.poll(() => copyText(page)).toBe('Body')
  const secondBytes = await download(page), second = await parse(page, secondBytes)
  expect(second.comments.map((comment) => ({ text: comment.text, ranges: comment.targets.map((target) => [target.start, target.end]) }))).toEqual([{ text: 'Keep label review', ranges: [[0, 0]] }])
  await upload(page, secondBytes, 'numbering-point-reopened-twice.docx')
  await expect.poll(() => copyText(page)).toBe('Body')
  await expect(page.locator('.document-surface canvas')).toHaveCount(1)
})

test('production clear formatting preserves document defaults through native DOCX saves and repeated import', async ({ page }, testInfo) => {
  const errors: string[] = [], workers: string[] = []
  page.on('pageerror', error => errors.push(error.message)); page.on('worker', worker => workers.push(worker.url()))
  await page.goto('/')
  const model = { kind: 'document', name: 'Clear formatting evidence', defaultStyle: { version: 1,
    text: { fontFamily: 'Georgia', eastAsiaFontFamily: '宋体', fontSize: 15, color: 'auto', bold: false, italic: false, underline: false, strike: false, highlight: '#ABCDEF' },
    paragraph: { align: 'right', indent: 0, lineHeight: 24, spacingRule: 'exact', spaceBefore: 0, spaceAfter: 12 },
  }, blocks: [{ id: 'formatted', type: 'heading', level: 3, style: { align: 'center', indent: 0, lineHeight: 2, spaceBefore: 8, spaceAfter: 20 }, metadata: {},
    runs: [{ id: 'body', text: 'Hello world', fontFamily: 'Arial', fontSize: 26, color: '#FF0000', bold: true, italic: true, underline: true, strike: true, highlight: '#FFFF00', verticalAlign: 'superscript' }],
  }] }
  const result = await officeWorker(page, { action: 'write', model: JSON.stringify(model) })
  expect(result.warnings).toEqual([])
  if (!result.bytes) throw new Error('Clear-format DOCX fixture has no bytes')
  const input = Buffer.from(result.bytes)
  await upload(page, input, 'clear-formatting-evidence.docx')
  const bridge = page.getByRole('textbox', { name: '文档文字输入', exact: true })
  await bridge.focus(); await bridge.press('Control+End')
  const clear = page.locator('[data-ribbon-item="doc-clear-format"]:visible').getByRole('button')
  await expect(clear).toBeEnabled(); await clear.click(); await expect(bridge).toBeFocused()
  const bold = page.locator('[data-ribbon-item="doc-bold"]:visible').getByRole('button')
  await expect(bold).toHaveAttribute('aria-pressed', 'false')
  await page.locator('[data-ribbon-item="doc-undo"]:visible').getByRole('button').click()
  await expect(bold).toHaveAttribute('aria-pressed', 'true')
  await page.locator('[data-ribbon-item="doc-redo"]:visible').getByRole('button').click()
  await expect(bold).toHaveAttribute('aria-pressed', 'false')
  const downloads: Array<{ sha256: string; bytes: number }> = []
  for (let iteration = 0; iteration < 2; iteration++) {
    const saved = await download(page), parsed = await parse(page, saved)
    const defaults = parsed.defaultStyle
    expect(defaults?.version).toBe(1)
    expect(defaults?.text?.fontFamilyAscii).toBe('Georgia')
    expect(defaults?.text?.fontFamilyEastAsia).toBe('宋体')
    expect(defaults?.text?.fontSizePt).toBe(15)
    expect(defaults?.text?.backgroundColor).toBe('#ABCDEF')
    expect(defaults?.paragraph?.align).toBe('right')
    expect(defaults?.paragraph?.spacing?.lineRule).toBe('exact')
    expect(defaults?.paragraph?.spacing?.lineTwips).toBe(360)
    expect(defaults?.paragraph?.spacing?.afterTwips).toBe(180)
    const paragraph = parsed.blocks.find(block => block.content === 'Hello world')
    const word = paragraph?.details.inlines.find(inline => inline.text === 'world')
    expect(word?.style?.fontFamilyAscii).toBe('Georgia'); expect(word?.style?.fontSizePt).toBe(15)
    expect(word?.style?.bold).toBe(false); expect(word?.style?.italic).toBe(false)
    expect(word?.style?.underline).toBe(false); expect(word?.style?.strike).toBe(false)
    expect(word?.style?.backgroundColor).toBe('#ABCDEF')
    const prefix = paragraph?.details.inlines.find(inline => inline.text === 'Hello ')
    expect(prefix?.style?.bold).toBe(true); expect(prefix?.style?.fontSizePt).toBe(26)
    const declaration = officePart(saved, 'word/styles.xml').split('</w:docDefaults>')[0]
    expect(declaration).toContain('w:ascii="Georgia"'); expect(declaration).toContain('w:eastAsia="宋体"')
    expect(declaration).toContain('w:fill="ABCDEF"')
    downloads.push({ sha256: createHash('sha256').update(saved).digest('hex'), bytes: saved.length })
    await upload(page, saved, `clear-formatting-reopened-${iteration + 1}.docx`)
    await bridge.focus(); await bridge.press('Control+End'); await clear.click(); await expect(bridge).toBeFocused()
  }
  expect(errors).toEqual([])
  expect(workers.some(url => /\/assets\/parser\.worker-/u.test(url))).toBe(true)
  expect(workers.some(url => /\/assets\/export\.worker-/u.test(url))).toBe(true)
  await testInfo.attach('clear-formatting-docx-evidence.json', { body: Buffer.from(JSON.stringify({ sourceSha256: createHash('sha256').update(input).digest('hex'), downloads, workers,
    limitation: 'Prepared production Worker and native DOCX reparse coverage; independent Office applications, system IME and large-file performance remain unverified.' }, null, 2)), contentType: 'application/json' })
})
