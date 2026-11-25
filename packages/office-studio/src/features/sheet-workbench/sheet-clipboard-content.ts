import { defaultCellStyle, type CellScalar, type SpreadsheetCell } from '../../domain/model.ts'
import { applyExcelCellView, toExcelCellView, type ExcelCellView } from '../../services/excelViewAdapter.ts'
import { assertSheetClipboardArea, assertSheetClipboardBytes, assertSheetClipboardSnapshot, sheetClipboardCellAt } from './sheet-clipboard-data.ts'
import type { SheetClipboardSnapshot } from './sheet-clipboard-types.ts'
import { sheetImageResourceFromDataUrl } from '../sheet-images/sheet-image-source.ts'
import { normalizeSheetCellImage } from '../sheet-images/sheet-image-model.ts'
import { SheetImageSourceNeedsPreparationError } from '../sheet-images/sheet-image-public.ts'

const HTML_ESCAPES: Readonly<Record<string, string>> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }
const CSS_UNIT_FACTORS: Readonly<Record<string, number>> = { px: 1, pt: 4 / 3, in: 96, cm: 96 / 2.54, mm: 96 / 25.4 }
const escapeHtml = (text: string): string => text.replace(/[&<>"']/g, character => HTML_ESCAPES[character]!)
const cellKey = (row: number, col: number): string => `${row}:${col}`

/** Quoted TSV carries tabs, quotes and line breaks inside one cell; a final record delimiter is not another row. */
export function parseSheetClipboardText(text: string): SheetClipboardSnapshot | undefined {
  if (text === '') return undefined
  assertSheetClipboardBytes(text)
  const rows: string[][] = [], row: string[] = []
  let field = '', quoted = false, closedQuote = false, lastWasRecord = false, widestRow = 1
  const appendField = (): void => {
    row.push(field); field = ''; closedQuote = false; widestRow = Math.max(widestRow, row.length)
    assertSheetClipboardArea(Math.max(1, rows.length + 1), widestRow)
  }
  const appendRow = (): void => {
    appendField(); rows.push(row.splice(0))
    assertSheetClipboardArea(rows.length, widestRow)
  }
  for (let index = 0; index < text.length; index++) {
    const character = text[index]!
    lastWasRecord = false
    if (quoted) {
      if (character === '"' && text[index + 1] === '"') { field += '"'; index++; continue }
      if (character === '"') { quoted = false; closedQuote = true; continue }
      if (character === '\r') { field += '\n'; if (text[index + 1] === '\n') index++; continue }
      field += character
      continue
    }
    if (character === '"' && field === '' && !closedQuote) { quoted = true; continue }
    if (closedQuote && character !== '\t' && character !== '\r' && character !== '\n') {
      throw new TypeError('剪贴板 TSV 的引号字段后包含无效字符。')
    }
    if (character === '\t') appendField()
    else if (character === '\r' || character === '\n') {
      appendRow(); lastWasRecord = true
      if (character === '\r' && text[index + 1] === '\n') index++
    } else field += character
  }
  if (quoted) throw new TypeError('剪贴板 TSV 的引号未闭合。')
  if (!lastWasRecord || row.length || field || closedQuote) appendRow()
  const columns = rows.reduce((width, current) => Math.max(width, current.length), 1)
  assertSheetClipboardArea(rows.length, columns)
  const cells: Record<string, SpreadsheetCell> = Object.create(null)
  for (const [rowIndex, values] of rows.entries()) for (const [col, value] of values.entries()) {
    const id = cellKey(rowIndex, col)
    cells[id] = { id, row: rowIndex, col, value, formula: null, display: value, dataType: value === '' ? 'blank' : 'text',
      cachedDataType: null, style: defaultCellStyle() }
  }
  return { version: 1, id: crypto.randomUUID(), copyType: 'copy', formulaOnly: false,
    rows: rows.map((_, index) => index), columns: Array.from({ length: columns }, (_, index) => index),
    cells, merges: [], columnWidths: Array.from({ length: columns }, () => 96), externalFormat: 'text' }
}

function quoteTsv(text: string): string { return /[\t\r\n"]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text }
function displayText(cell: SpreadsheetCell | undefined, formulaOnly: boolean): string {
  if (!cell) return ''
  return formulaOnly ? cell.formula ?? '' : cell.display || (cell.value === null ? '' : String(cell.value))
}
function fontCss(cell: SpreadsheetCell, run?: Record<string, unknown>): string {
  const font = cell.style.font
  const nested = run?.font && typeof run.font === 'object' && !Array.isArray(run.font) ? run.font as Record<string, unknown> : undefined
  const mark = (key: string, fallback: unknown): unknown => run?.[key] ?? nested?.[key] ?? fallback
  return [`font-family:${String(run?.fontFamily ?? nested?.family ?? font.family)}`, `font-size:${Number(run?.fontSize ?? nested?.size ?? font.size)}pt`,
    `font-weight:${mark('bold', font.bold) ? 'bold' : 'normal'}`, `font-style:${mark('italic', font.italic) ? 'italic' : 'normal'}`,
    `color:${String(mark('color', font.color) ?? 'inherit')}`,
    `text-decoration:${[mark('underline', font.underline) ? 'underline' : '', mark('strike', font.strike) ? 'line-through' : ''].filter(Boolean).join(' ') || 'none'}`].join(';')
}
function cellCss(cell: SpreadsheetCell): string {
  const style = cell.style
  return [fontCss(cell), `background-color:${style.fill ?? 'transparent'}`, `text-align:${style.horizontal === 'general' ? 'initial' : style.horizontal}`,
    `vertical-align:${style.vertical === 'middle' ? 'middle' : style.vertical}`, `white-space:${style.wrap ? 'pre-wrap' : 'pre'}`,
    `mso-number-format:${JSON.stringify(style.numberFormat)}`,
    ...Object.entries(style.border ?? {}).flatMap(([side, border]) => {
      if (!['top', 'right', 'bottom', 'left'].includes(side) || !border || typeof border !== 'object') return []
      const value = border as Record<string, unknown>
      const cssStyle = ['double', 'dashed', 'dotted'].includes(String(value.style)) ? String(value.style) : 'solid'
      return [`border-${side}:${value.style === 'double' ? 3 : value.style === 'thick' ? 3 : value.style === 'medium' ? 2 : 1}px ${cssStyle} ${String(value.color ?? '#000000')}`]
    })].join(';')
}
function cellHtml(cell: SpreadsheetCell | undefined, formulaOnly: boolean): string {
  if (!cell) return ''
  if (formulaOnly || cell.formula) return escapeHtml(displayText(cell, formulaOnly)).replace(/\r\n?|\n/g, '<br>')
  const runs = cell.richText
  const text = runs?.length && runs.every(run => typeof run.text === 'string')
    ? runs.map(run => {
      const text = `<span style="${escapeHtml(fontCss(cell, run))}">${escapeHtml(String(run.text)).replace(/\r\n?|\n/g, '<br>')}</span>`
      return typeof run.hyperlink === 'string' && safeHref(run.hyperlink) ? `<a href="${escapeHtml(run.hyperlink)}">${text}</a>` : text
    }).join('')
    : escapeHtml(displayText(cell, false)).replace(/\r\n?|\n/g, '<br>')
  const linked = cell.hyperlink ? `<a href="${escapeHtml(cell.hyperlink)}">${text}</a>` : text
  const image = cell.image
  return image ? `${linked}<img src="${escapeHtml(image.resource.source)}" width="${image.resource.width}" height="${image.resource.height}" alt="${escapeHtml(image.altText)}" title="${escapeHtml(image.title)}">` : linked
}

/** The plain flavour contains evaluated values; the internal HTML token selects the complete snapshot. */
export function sheetClipboardContent(snapshot: SheetClipboardSnapshot): { plain: string; html: string } {
  assertSheetClipboardSnapshot(snapshot)
  const covered = new Set<string>(), anchors = new Map(snapshot.merges.map(merge => [cellKey(merge.startRow, merge.startCol), merge]))
  for (const merge of snapshot.merges) for (let row = merge.startRow; row <= merge.endRow; row++) for (let col = merge.startCol; col <= merge.endCol; col++) {
    if (row !== merge.startRow || col !== merge.startCol) covered.add(cellKey(row, col))
  }
  const plainRows: string[] = [], htmlRows: string[] = []
  for (let row = 0; row < snapshot.rows.length; row++) {
    const plain: string[] = [], html: string[] = []
    for (let col = 0; col < snapshot.columns.length; col++) {
      const key = cellKey(row, col), cell = sheetClipboardCellAt(snapshot, row, col), merge = anchors.get(key)
      plain.push(quoteTsv(displayText(covered.has(key) ? undefined : cell, snapshot.formulaOnly)))
      if (covered.has(key)) continue
      const formula = cell?.formula ? ` data-office-formula="${escapeHtml(cell.formula)}"` : ''
      const value = cell ? ` data-office-value="${escapeHtml(JSON.stringify(cell.value))}" data-office-value-type="${escapeHtml(cell.dataType ?? '')}"` : ''
      html.push(`<td${merge ? ` rowspan="${merge.endRow - merge.startRow + 1}" colspan="${merge.endCol - merge.startCol + 1}"` : ''}${formula}${value}${cell ? ` style="${escapeHtml(cellCss(cell))}"` : ''}>${cellHtml(cell, snapshot.formulaOnly)}</td>`)
    }
    plainRows.push(plain.join('\t')); htmlRows.push(`<tr>${html.join('')}</tr>`)
  }
  const plain = plainRows.join('\r\n'), html = `<table data-office-clipboard-id="${escapeHtml(snapshot.id)}"><colgroup>${snapshot.columnWidths.map(width => `<col style="width:${width}px">`).join('')}</colgroup><tbody>${htmlRows.join('')}</tbody></table>`
  assertSheetClipboardBytes(plain); assertSheetClipboardBytes(html)
  return { plain, html }
}

export function sheetClipboardIdFromHtml(html: string): string | undefined {
  const matched = /\bdata-office-clipboard-id\s*=\s*(?:"([A-Za-z0-9._:-]+)"|'([A-Za-z0-9._:-]+)')/i.exec(html)
  return matched?.[1] ?? matched?.[2]
}

function cssDeclarations(text: string): Record<string, string> {
  const values: Record<string, string> = Object.create(null)
  let property = '', quote = '', depth = 0
  const append = (): void => {
    const separator = property.indexOf(':')
    if (separator > 0) values[property.slice(0, separator).trim().toLowerCase()] = property.slice(separator + 1).trim().replace(/\s*!important\s*$/i, '')
    property = ''
  }
  for (let index = 0; index < text.length; index++) {
    const character = text[index]!
    if (character === '\\' && index + 1 < text.length) { property += character + text[++index]!; continue }
    if (quote) { property += character; if (character === quote) quote = ''; continue }
    if (character === '"' || character === "'") quote = character
    else if (character === '(') depth++
    else if (character === ')') depth = Math.max(0, depth - 1)
    else if (character === ';' && depth === 0) { append(); continue }
    property += character
  }
  append()
  return values
}
function cssPixels(value: string | null | undefined): number | undefined {
  if (!value) return undefined
  const match = /^\s*(\d+(?:\.\d+)?)(px|pt|in|cm|mm)?\s*$/i.exec(value)
  if (!match) return undefined
  return Number(match[1]) * (CSS_UNIT_FACTORS[match[2]?.toLowerCase() ?? 'px'] ?? 1)
}
function unquoteCss(text: string): string {
  const raw = /^(["']).*\1$/.test(text) ? text.slice(1, -1) : text
  return raw.replace(/\\(["'\\;])/g, '$1')
}
function safeHref(value: string): string | null { return /^(?:https?:|mailto:|tel:|#)/i.test(value.trim()) ? value.trim() : null }
function styleView(properties: Record<string, string>, inherited: ExcelCellView = { value: '' }): ExcelCellView {
  const result = { ...inherited }
  const fontSize = cssPixels(properties['font-size'])
  if (fontSize !== undefined) result.fontSize = fontSize * .75
  if (properties['font-family']) result.fontFamily = unquoteCss(properties['font-family'])
  if (properties['font-weight']) result.bold = ['bold', 'bolder'].includes(properties['font-weight']) || Number(properties['font-weight']) >= 600
  if (properties['font-style']) result.italic = properties['font-style'] === 'italic'
  if (properties.color) result.color = properties.color
  if (properties['background-color']) result.background = properties['background-color'] === 'transparent' ? null : properties['background-color']
  if (properties['text-decoration'] || properties['text-decoration-line']) {
    const decoration = properties['text-decoration'] ?? properties['text-decoration-line']!
    result.underline = decoration.includes('underline'); result.strike = decoration.includes('line-through')
  }
  const align = properties['text-align'], vertical = properties['vertical-align']
  if (align === 'left' || align === 'center' || align === 'right' || align === 'justify') result.align = align
  if (vertical === 'top' || vertical === 'middle' || vertical === 'bottom') result.vertical = vertical
  if (properties['white-space']) {
    result.wrap = ['normal', 'pre-wrap', 'pre-line', 'break-spaces'].includes(properties['white-space'])
    result.wrapStrategy = result.wrap ? 'wrap' : 'overflow'
  }
  if (properties['mso-number-format']) result.rawNumberFormat = unquoteCss(properties['mso-number-format'])
  const borders: Record<string, unknown> = {}
  for (const side of ['top', 'right', 'bottom', 'left']) {
    const border = properties[`border-${side}`] ?? properties.border
    if (!border) continue
    if (/\b(?:none|hidden)\b|^0(?:px|pt)?\b/i.test(border)) { borders[side] = null; continue }
    const matched = /^(\S+)\s+(solid|double|dotted|dashed)\s+(.+)$/i.exec(border)
    if (matched) borders[side] = { style: matched[2] === 'solid' ? (cssPixels(matched[1]) ?? 1) > 2 ? 'thick' : (cssPixels(matched[1]) ?? 1) > 1 ? 'medium' : 'thin' : matched[2], color: matched[3] }
  }
  if (Object.keys(borders).length) result.border = borders
  return result
}
function inlineHtmlRuns(root: Element, base: ExcelCellView, styles: Map<string, Record<string, string>>): { text: string; runs: Record<string, unknown>[]; hyperlink: string | null } {
  const runs: Record<string, unknown>[] = []
  let hyperlink: string | null = null
  const walk = (node: Node, inherited: ExcelCellView, href?: string | null): void => {
    if (node.nodeType === 3) {
      const text = node.nodeValue?.replace(/\r\n?/g, '\n') ?? ''
      if (text) runs.push({ text, fontFamily: inherited.fontFamily, fontSize: inherited.fontSize, bold: inherited.bold, italic: inherited.italic,
        underline: inherited.underline, strike: inherited.strike, color: inherited.color, ...(href ? { hyperlink: href } : {}) })
      return
    }
    if (node.nodeType !== 1) return
    const element = node as Element, tag = element.tagName.toLowerCase()
    if (['script', 'style', 'iframe', 'object', 'embed', 'meta', 'link', 'title', 'base'].includes(tag)) return
    if (tag === 'img') return
    if (tag === 'br') { runs.push({ text: '\n' }); return }
    let marks = inherited
    for (const name of element.getAttribute('class')?.split(/\s+/) ?? []) if (styles.has(name)) marks = styleView(styles.get(name)!, marks)
    marks = styleView(cssDeclarations(element.getAttribute('style') ?? ''), marks)
    if (tag === 'b' || tag === 'strong') marks = { ...marks, bold: true }
    if (tag === 'i' || tag === 'em') marks = { ...marks, italic: true }
    if (tag === 'u') marks = { ...marks, underline: true }
    if (tag === 's' || tag === 'strike' || tag === 'del') marks = { ...marks, strike: true }
    const link = tag === 'a' ? safeHref(element.getAttribute('href') ?? '') : href
    if (link) hyperlink = link
    if (['p', 'div'].includes(tag) && runs.length && runs.at(-1)?.text !== '\n') runs.push({ text: '\n' })
    for (const child of Array.from(element.childNodes)) walk(child, marks, link)
  }
  walk(root, base)
  return { text: runs.map(run => String(run.text)).join(''), runs, hyperlink }
}
function htmlTypedValue(element: Element): { value?: CellScalar; dataType?: string; formula?: string } {
  const ownValue = element.getAttribute('data-office-value'), numeric = element.getAttribute('x:num')
  if (ownValue !== null) {
    let parsed: unknown
    try { parsed = JSON.parse(ownValue) } catch { throw new TypeError('剪贴板 HTML 的单元格原始值无效。') }
    if (parsed !== null && typeof parsed !== 'string' && typeof parsed !== 'number' && typeof parsed !== 'boolean'
      || typeof parsed === 'number' && !Number.isFinite(parsed)) throw new TypeError('剪贴板 HTML 的单元格类型无效。')
    return { value: parsed, dataType: element.getAttribute('data-office-value-type') ?? undefined,
      formula: element.getAttribute('data-office-formula') ?? undefined }
  }
  if (numeric !== null && Number.isFinite(Number(numeric))) return { value: Number(numeric), dataType: 'number' }
  return { formula: element.getAttribute('data-office-formula') ?? undefined }
}

/** HTML is parsed in an inert document. The returned cells, runs, merges and widths are the model input. */
export function parseSheetClipboardHtml(html: string): SheetClipboardSnapshot | undefined {
  if (!html.trim()) return undefined
  assertSheetClipboardBytes(html)
  if (typeof DOMParser === 'undefined') throw new Error('当前环境缺少 HTML 剪贴板解析器。')
  const parsed = new DOMParser().parseFromString(html, 'text/html'), table = parsed.querySelector('table')
  if (parsed.querySelector('svg,canvas,video,audio,object,embed')) throw new Error('剪贴板包含尚未接入的绘图或附件资源，未修改目标区域。')
  if (table && Array.from(parsed.querySelectorAll('img')).some(image => !table.contains(image))) throw new Error('剪贴板图片位于表格区域外，无法确定其单元格位置。')
  const styles = new Map<string, Record<string, string>>()
  for (const sheet of Array.from(parsed.querySelectorAll('style'))) {
    for (const match of (sheet.textContent ?? '').matchAll(/\.([\w-]+)\s*\{([^{}]+)\}/g)) styles.set(match[1]!, cssDeclarations(match[2]!))
  }
  const rows = table ? Array.from((table as HTMLTableElement).rows) : []
  const cells: Record<string, SpreadsheetCell> = Object.create(null), merges: SheetClipboardSnapshot['merges'] = [], occupied = new Set<string>()
  let rowCount = Math.max(1, rows.length), columnCount = 1
  const append = (element: Element, row: number, col: number): void => {
    let sourceView: ExcelCellView = toExcelCellView({ id: cellKey(row, col), row, col, value: null, formula: null, display: '', style: defaultCellStyle() })
    for (const name of element.getAttribute('class')?.split(/\s+/) ?? []) if (styles.has(name)) sourceView = styleView(styles.get(name)!, sourceView)
    sourceView = styleView(cssDeclarations(element.getAttribute('style') ?? ''), sourceView)
    const inline = inlineHtmlRuns(element, sourceView, styles), typed = htmlTypedValue(element)
    sourceView.value = typed.value === undefined ? inline.text : typed.value === null ? '' : String(typed.value)
    sourceView.rawValue = typed.value === undefined ? inline.text : typed.value
    sourceView.dataType = typed.dataType ?? 'text'; sourceView.formula = typed.formula
    sourceView.richText = !typed.formula && (inline.runs.length > 1 || inline.hyperlink || element.querySelector('b,strong,i,em,u,s,strike,span,a')) ? inline.runs : null
    sourceView.hyperlink = inline.hyperlink
    sourceView.displayValue = inline.text
    const cell = applyExcelCellView({ row, col }, sourceView)
    const images = Array.from(element.querySelectorAll('img'))
    if (images.length > 1) throw new Error('一个剪贴板单元格包含多张图片，当前图片正文布局尚未覆盖，未修改目标区域。')
    const image = images[0]
    if (image) {
      const source = image.getAttribute('src') ?? ''
      if (!/^data:/i.test(source)) throw new SheetImageSourceNeedsPreparationError()
      const resource = sheetImageResourceFromDataUrl(source, image.getAttribute('title') || 'clipboard-image')
      cell.image = normalizeSheetCellImage({ id: crypto.randomUUID(), resource, title: image.getAttribute('title') ?? '', altText: image.getAttribute('alt') ?? '',
        metadata: { clipboardHtmlImage: { width: image.getAttribute('width'), height: image.getAttribute('height'), style: image.getAttribute('style') ?? '' } } })
    }
    cell.metadata = { ...cell.metadata, clipboardHtml: element.innerHTML }
    cell.style.metadata = { ...cell.style.metadata, clipboardCss: element.getAttribute('style') ?? '' }
    cells[cell.id] = cell
  }
  if (!table) append(parsed.body, 0, 0)
  else for (const [row, element] of rows.entries()) {
    let col = 0
    for (const cell of Array.from(element.cells)) {
      while (occupied.has(cellKey(row, col))) col++
      const rowSpan = cell.rowSpan === 0 ? rows.length - row : cell.rowSpan, colSpan = cell.colSpan
      rowCount = Math.max(rowCount, row + rowSpan); columnCount = Math.max(columnCount, col + colSpan)
      assertSheetClipboardArea(rowCount, columnCount)
      append(cell, row, col)
      if (rowSpan > 1 || colSpan > 1) merges.push({ startRow: row, endRow: row + rowSpan - 1, startCol: col, endCol: col + colSpan - 1 })
      for (let r = row; r < row + rowSpan; r++) for (let c = col; c < col + colSpan; c++) occupied.add(cellKey(r, c))
      col += colSpan
    }
  }
  const widths = Array.from({ length: columnCount }, () => 96)
  let column = 0
  for (const element of Array.from(table?.querySelectorAll('col') ?? [])) {
    const width = cssPixels(cssDeclarations(element.getAttribute('style') ?? '').width ?? element.getAttribute('width'))
    const span = Math.max(1, Number(element.getAttribute('span')) || 1)
    for (let offset = 0; offset < span && column < columnCount; offset++, column++) if (width !== undefined) widths[column] = width
  }
  const snapshot: SheetClipboardSnapshot = { version: 1, id: crypto.randomUUID(), copyType: 'copy', formulaOnly: false,
    rows: Array.from({ length: rowCount }, (_, index) => index), columns: Array.from({ length: columnCount }, (_, index) => index),
    cells, merges, columnWidths: widths, externalFormat: 'html' }
  assertSheetClipboardSnapshot(snapshot)
  return snapshot
}
