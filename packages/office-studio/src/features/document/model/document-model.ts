import type { DocumentBlock, DocumentModel, DocumentParagraphStyle, InlineRun } from '../../../domain/model'

export interface DocumentRun extends InlineRun { highlight?: string; href?: string }
export interface TextTarget { blockId: string; row?: number; col?: number; cellBlockId?: string }
export interface DocumentPoint extends TextTarget { offset: number }
export interface DocumentSelection { anchor: DocumentPoint; focus: DocumentPoint }
export interface DocumentCell {
  id: string
  blocks: DocumentBlock[]
  style: Record<string, unknown>
  rowSpan: number
  colSpan: number
}
export interface DocumentTable { rows: Array<{ id: string; cells: DocumentCell[] }>; columnWidths?: number[]; headerRowCount?: number; style?: Record<string, unknown> }
export interface DocumentImage { src: string; alt: string; width: number; height: number }
export type RunPatch = Partial<Omit<DocumentRun, 'id' | 'text'>>

export type DocumentNamedParagraphStyle = 'normal' | 'title' | 'subtitle' | 'heading1' | 'heading2' | 'heading3' | 'heading4' | 'heading5'

export const DOCUMENT_NAMED_PARAGRAPH_STYLES: Record<DocumentNamedParagraphStyle, {
  fontSize: number
  bold: boolean
  color?: string
  level: number
  spaceBefore: number
  spaceAfter: number
}> = {
  normal: { fontSize: 11, bold: false, level: 0, spaceBefore: 0, spaceAfter: 16 },
  title: { fontSize: 26, bold: true, level: 0, spaceBefore: 0, spaceAfter: 9.3 },
  subtitle: { fontSize: 15, bold: false, color: '#999999', level: 0, spaceBefore: 0, spaceAfter: 16 },
  heading1: { fontSize: 20, bold: true, level: 1, spaceBefore: 26.7, spaceAfter: 13.3 },
  heading2: { fontSize: 18, bold: true, level: 2, spaceBefore: 24, spaceAfter: 13.3 },
  heading3: { fontSize: 16, bold: true, level: 3, spaceBefore: 21.3, spaceAfter: 13.3 },
  heading4: { fontSize: 14, bold: true, level: 4, spaceBefore: 18.7, spaceAfter: 10.7 },
  heading5: { fontSize: 12, bold: true, level: 5, spaceBefore: 16, spaceAfter: 10.7 },
}

export function applyNamedParagraphStyle(block: DocumentBlock, namedStyle: DocumentNamedParagraphStyle): void {
  const style = DOCUMENT_NAMED_PARAGRAPH_STYLES[namedStyle]
  block.metadata = { ...block.metadata, namedStyleType: namedStyle }
  block.type = style.level > 0 ? 'heading' : 'paragraph'
  block.level = style.level
  block.style = {
    ...block.style,
    indent: 0,
    lineHeight: 1.5,
    spaceBefore: style.spaceBefore,
    spaceAfter: style.spaceAfter,
  }
  block.runs = block.runs.map((run) => ({
    ...run,
    fontSize: style.fontSize,
    bold: style.bold,
    ...(style.color ? { color: style.color } : namedStyle === 'normal' ? { color: '#1f2937' } : {}),
  }))
}

export function applyRunPatch(run: DocumentRun, patch: RunPatch): DocumentRun {
  const updated = { ...run, ...patch }
  if (patch.verticalAlign === 'baseline') {
    delete updated.verticalAlign
    if (updated.metadata?.verticalAlign !== undefined) {
      const metadata = { ...updated.metadata }
      delete metadata.verticalAlign
      if (Object.keys(metadata).length) updated.metadata = metadata
      else delete updated.metadata
    }
  }
  return updated
}

let sequence = 0
export const documentId = (prefix: string): string => `${prefix}-${Date.now().toString(36)}-${(++sequence).toString(36)}`
export const cloneDocument = (value: DocumentModel): DocumentModel => JSON.parse(JSON.stringify(value)) as DocumentModel
export const defaultParagraphStyle = (): DocumentParagraphStyle => ({ align: 'left', indent: 0, lineHeight: 1.5, spaceBefore: 0, spaceAfter: 8 })
export const createRun = (text = '', patch: RunPatch = {}): DocumentRun => applyRunPatch({ id: documentId('run'), text, bold: false, italic: false, underline: false, strike: false, color: '#1f2937', fontFamily: 'Arial', fontSize: 11, ...patch }, patch)
export const createParagraph = (text = '', patch: RunPatch = {}): DocumentBlock => ({ id: documentId('paragraph'), type: 'paragraph', level: 0, runs: [createRun(text, patch)], style: defaultParagraphStyle(), metadata: {} })
export const createChecklistParagraph = (text = '', checked = false, listId = documentId('checklist')): DocumentBlock => ({ ...createParagraph(text), type: 'checklist', checklist: { listId, checked } })
export const targetKey = (target: TextTarget): string => target.row == null ? target.blockId : `${target.blockId}/r${target.row}/c${target.col ?? 0}/${target.cellBlockId ?? ''}`
export const sameTarget = (a: TextTarget, b: TextTarget): boolean => targetKey(a) === targetKey(b)
export const collapsedSelection = (point: DocumentPoint): DocumentSelection => ({ anchor: { ...point }, focus: { ...point } })
export const selectionCollapsed = (selection: DocumentSelection): boolean => sameTarget(selection.anchor, selection.focus) && selection.anchor.offset === selection.focus.offset
export const runText = (runs: readonly InlineRun[]): string => runs.map((run) => run.text).join('')

export function tableForBlock(block: DocumentBlock): DocumentTable | undefined {
  const table = block.metadata.table
  if (!table || typeof table !== 'object' || !('rows' in table) || !Array.isArray(table.rows)) return undefined
  return table as DocumentTable
}

export function imageForBlock(block: DocumentBlock): DocumentImage | undefined {
  const image = block.metadata.image
  if (!image || typeof image !== 'object' || !('src' in image) || typeof image.src !== 'string') return undefined
  return image as DocumentImage
}

export function textTargets(model: DocumentModel): TextTarget[] {
  const targets: TextTarget[] = []
  for (const block of model.blocks) {
    const table = tableForBlock(block)
    if (block.type === 'table' && table) {
      table.rows.forEach((row, rowIndex) => row.cells.forEach((cell, colIndex) => cell.blocks.forEach((cellBlock) => {
        if (!['table', 'image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak'].includes(cellBlock.type)) targets.push({ blockId: block.id, row: rowIndex, col: colIndex, cellBlockId: cellBlock.id })
      })))
    } else if (!['image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak'].includes(block.type)) targets.push({ blockId: block.id })
  }
  return targets
}

export function runsForTarget(model: DocumentModel, target: TextTarget): DocumentRun[] {
  const block = model.blocks.find((candidate) => candidate.id === target.blockId)
  if (!block) return []
  if (target.row != null) {
    const cell = tableForBlock(block)?.rows[target.row]?.cells[target.col ?? 0]
    return (cell?.blocks.find((candidate) => candidate.id === target.cellBlockId)?.runs ?? cell?.blocks[0]?.runs ?? []) as DocumentRun[]
  }
  return block.runs as DocumentRun[]
}

export function setTargetRuns(model: DocumentModel, target: TextTarget, runs: DocumentRun[]): void {
  const block = model.blocks.find((candidate) => candidate.id === target.blockId)
  if (!block) return
  const nextRuns = compactRuns(runs)
  if (target.row != null) {
    const cell = tableForBlock(block)?.rows[target.row]?.cells[target.col ?? 0]
    const paragraph = cell?.blocks.find((candidate) => candidate.id === target.cellBlockId) ?? cell?.blocks[0]
    if (paragraph) paragraph.runs = nextRuns
  } else block.runs = nextRuns
  delete block.metadata.html
}

export function compactRuns(runs: DocumentRun[]): DocumentRun[] {
  const result: DocumentRun[] = []
  for (const source of runs) {
    if (!source.text) continue
    const normalized = source.verticalAlign === 'baseline' ? applyRunPatch(source, { verticalAlign: 'baseline' }) : source
    const last = result.at(-1)
    const { id: _id, text: _text, ...marks } = normalized
    if (last) {
      const { id: _lastId, text: _lastText, ...lastMarks } = last
      if (JSON.stringify(marks) === JSON.stringify(lastMarks)) { last.text += normalized.text; continue }
    }
    result.push({ ...normalized, id: documentId('run') })
  }
  return result.length ? result : [createRun()]
}

export function sliceRuns(runs: readonly DocumentRun[], start: number, end = Number.POSITIVE_INFINITY): DocumentRun[] {
  const result: DocumentRun[] = []
  let offset = 0
  for (const run of runs) {
    const from = Math.max(0, start - offset)
    const to = Math.min(run.text.length, end - offset)
    if (to > from) result.push({ ...run, id: documentId('run'), text: run.text.slice(from, to) })
    offset += run.text.length
    if (offset >= end) break
  }
  return result
}

export function clampPoint(model: DocumentModel, point: DocumentPoint): DocumentPoint {
  const targets = textTargets(model)
  const block = point.row == null ? undefined : model.blocks.find((candidate) => candidate.id === point.blockId)
  const table = block && tableForBlock(block)
  const row = table ? Math.max(0, Math.min(point.row ?? 0, table.rows.length - 1)) : undefined
  const col = row == null ? undefined : Math.max(0, Math.min(point.col ?? 0, (table?.rows[row]?.cells.length ?? 1) - 1))
  // Undo/redo can replace a cell paragraph id while retaining its logical table coordinate.
  const target = targets.find((candidate) => sameTarget(candidate, point))
    ?? targets.find((candidate) => candidate.blockId === point.blockId && candidate.row === row && candidate.col === col)
    ?? targets.find((candidate) => candidate.blockId === point.blockId)
    ?? targets[0]
  return target ? { ...target, offset: Math.max(0, Math.min(runText(runsForTarget(model, target)).length, point.offset)) } : { blockId: model.blocks[0]?.id ?? '', offset: 0 }
}

export function orderedSelection(model: DocumentModel, selection: DocumentSelection): { start: DocumentPoint; end: DocumentPoint } {
  const anchor = clampPoint(model, selection.anchor)
  const focus = clampPoint(model, selection.focus)
  const targets = textTargets(model)
  const a = targets.findIndex((target) => sameTarget(target, anchor))
  const b = targets.findIndex((target) => sameTarget(target, focus))
  return a < b || (a === b && anchor.offset <= focus.offset) ? { start: anchor, end: focus } : { start: focus, end: anchor }
}

export function marksAtPoint(model: DocumentModel, point: DocumentPoint): RunPatch {
  const runs = runsForTarget(model, point)
  let offset = 0
  const run = runs.find((candidate) => { offset += candidate.text.length; return point.offset <= offset }) ?? runs.at(-1) ?? createRun()
  const { id: _id, text: _text, ...patch } = run
  return patch
}

export function selectedPlainText(model: DocumentModel, selection: DocumentSelection): string {
  const { start, end } = orderedSelection(model, selection)
  const targets = textTargets(model)
  const first = targets.findIndex((target) => sameTarget(target, start))
  const last = targets.findIndex((target) => sameTarget(target, end))
  return targets.slice(first, last + 1).map((target, index) => {
    const text = runText(runsForTarget(model, target))
    return text.slice(index === 0 ? start.offset : 0, first + index === last ? end.offset : undefined)
  }).join('\n')
}

export function normalizeDocument(model: DocumentModel): DocumentModel {
  const result = cloneDocument(model)
  result.blocks = result.blocks.flatMap((block) => {
    const html = block.metadata.html
    if (typeof html !== 'string') return [block]
    const parsed = blocksFromHtml(html)
    if (!parsed.length) return [block]
    const first = parsed[0]
    return [{ ...first, id: block.id, type: block.type === 'paragraph' ? first.type : block.type, level: block.level || first.level, style: { ...first.style, ...block.style }, metadata: { ...block.metadata, ...first.metadata, sourceHtml: html, html: undefined } }, ...parsed.slice(1)]
  })
  for (const block of result.blocks) {
    if (block.type === 'table' && !tableForBlock(block)) block.metadata.table = createTable(1, 1)
    if (block.type === 'checklist') {
      block.checklist = {
        listId: typeof block.checklist?.listId === 'string' && block.checklist.listId ? block.checklist.listId : documentId('checklist'),
        checked: block.checklist?.checked === true,
      }
    } else delete block.checklist
    if (block.type === 'image' && !imageForBlock(block)) {
      const src = typeof block.metadata.src === 'string' ? block.metadata.src : typeof block.metadata.source === 'string' ? block.metadata.source : ''
      block.metadata.image = { src, alt: String(block.metadata.alt ?? ''), width: Number(block.metadata.width ?? 0), height: Number(block.metadata.height ?? 0) } satisfies DocumentImage
    }
  }
  if (!textTargets(result).length) result.blocks.push(createParagraph())
  result.wordCount = textTargets(result).reduce((count, target) => count + runText(runsForTarget(result, target)).replace(/\s/g, '').length, 0)
  return result
}

export function createTable(rows: number, cols: number): DocumentTable {
  return { rows: Array.from({ length: rows }, () => ({ id: documentId('row'), cells: Array.from({ length: cols }, () => ({ id: documentId('cell'), blocks: [createParagraph()], style: { fill: '#ffffff' }, rowSpan: 1, colSpan: 1 })) })), columnWidths: [], headerRowCount: 0 }
}

/** HTML is accepted at import/clipboard boundaries only; it is never the editor state. */
export function blocksFromHtml(html: string): DocumentBlock[] {
  const root = document.createElement('div')
  root.innerHTML = html
  root.querySelectorAll('script,style,iframe,object,embed,meta,link,title,base').forEach((node) => node.remove())
  const blocks: DocumentBlock[] = []
  const append = (element: HTMLElement) => {
    const tag = element.tagName
    if (tag === 'PRE' && element.dataset.officeCodeBlock === 'v1') {
      const source = element.querySelector('code')?.textContent ?? element.textContent ?? ''
      const languageValue = element.dataset.officeLanguage ?? 'plain'
      const language = ['plain', 'javascript', 'typescript', 'json', 'css', 'html', 'python', 'sql', 'rust'].includes(languageValue) ? languageValue : 'plain'
      const block = createParagraph(source, { fontFamily: 'Consolas', fontSize: 10 })
      block.metadata.codeBlock = { source, language, lineNumbers: element.dataset.officeLineNumbers === 'true' }
      block.style = { ...block.style, align: 'left', indent: 0, lineHeight: 1.35, spaceBefore: 8, spaceAfter: 8 }
      blocks.push(block)
      return
    }
    if (tag === 'DIV' && element.dataset.officeLatex === 'v1' && element.dataset.officeLatexMode === 'block') {
      const latex = element.dataset.officeLatexSource ?? element.textContent ?? ''
      const run = createRun('\uFFFC', { fontFamily: 'Cambria Math' })
      run.metadata = { ...run.metadata, latexFormula: { latex, mode: 'block', version: 1 } }
      const block = createParagraph()
      block.runs = [run]
      block.style = { ...block.style, align: 'center', spaceBefore: 8, spaceAfter: 8 }
      block.metadata.latexFormulaBlock = true
      blocks.push(block)
      return
    }
    if (tag === 'UL' || tag === 'OL') {
      const checklist = tag === 'UL' && element.dataset.officeListType === 'checklist'
      const listId = checklist ? element.dataset.officeListId || documentId('checklist') : undefined
      Array.from(element.children).forEach((child) => {
        const block = createParagraph()
        const item = child as HTMLElement
        const checkbox = item.querySelector<HTMLInputElement>('input[type="checkbox"]')
        const isChecklistItem = checklist || checkbox !== null
        block.type = isChecklistItem ? 'checklist' : tag === 'UL' ? 'bullet' : 'numbered'
        if (isChecklistItem) block.checklist = {
          listId: listId ?? documentId('checklist'),
          checked: item.dataset.officeChecked === 'true' || checkbox?.checked === true,
        }
        block.runs = inlineRunsFromElement(item)
        blocks.push(block)
      })
      return
    }
    const block = createParagraph()
    if (/^H[1-6]$/.test(tag)) { block.type = 'heading'; block.level = Number(tag.slice(1)) }
    if (tag === 'BLOCKQUOTE') block.type = 'quote'
    if (tag === 'HR') {
      const sectionType = element.dataset.officeSectionBreak
      if (sectionType === 'unspecified' || sectionType === 'nextPage' || sectionType === 'continuous' || sectionType === 'nextColumn' || sectionType === 'evenPage' || sectionType === 'oddPage') {
        block.type = 'sectionBreak'
        block.sectionBreak = { type: sectionType }
      } else if (element.dataset.officeColumnBreak) block.type = 'columnBreak'
      else block.type = element.dataset.officePageBreak ? 'pageBreak' : 'divider'
    }
    if (tag === 'TABLE') {
      block.type = 'table'
      const rows = Array.from((element as HTMLTableElement).rows).map((row) => ({ id: documentId('row'), cells: Array.from(row.cells).map((cell) => {
        const paragraph = createParagraph()
        paragraph.runs = inlineRunsFromElement(cell)
        paragraph.style.align = textAlign(cell.style.textAlign)
        return { id: documentId('cell'), blocks: [paragraph], style: { fill: cell.style.backgroundColor || '#ffffff' }, rowSpan: cell.rowSpan, colSpan: cell.colSpan }
      }) }))
      block.metadata.table = { rows, columnWidths: [], headerRowCount: (element as HTMLTableElement).tHead?.rows.length ?? 0 } satisfies DocumentTable
    }
    const image = tag === 'IMG' ? element as HTMLImageElement : element.querySelector<HTMLImageElement>('img')
    if (image && (!element.textContent?.trim() || tag === 'IMG')) {
      block.type = 'image'
      block.metadata.image = { src: safeUrl(image.getAttribute('src') ?? ''), alt: image.alt, width: Number(image.getAttribute('width')) || 0, height: Number(image.getAttribute('height')) || 0 } satisfies DocumentImage
    }
    block.runs = inlineRunsFromElement(element, block.type === 'heading' ? { fontSize: block.level === 1 ? 24 : block.level === 2 ? 18 : 14, bold: true } : {})
    block.style.align = textAlign(element.style.textAlign)
    const lineHeight = Number.parseFloat(element.style.lineHeight)
    if (lineHeight > 0 && lineHeight < 5) block.style.lineHeight = lineHeight
    block.style.spaceBefore = Number.parseFloat(element.style.marginTop) || 0
    block.style.spaceAfter = Number.parseFloat(element.style.marginBottom) || 8
    blocks.push(block)
  }
  root.childNodes.forEach((node) => {
    if (node.nodeType === Node.TEXT_NODE && node.textContent?.trim()) blocks.push(createParagraph(node.textContent))
    if (node instanceof HTMLElement) append(node)
  })
  return blocks
}

function textAlign(value: string): DocumentParagraphStyle['align'] { return value === 'center' || value === 'right' || value === 'justify' ? value : 'left' }
export function safeUrl(value: string): string { return /^(https?:|mailto:|#|blob:|data:image\/(png|jpeg|gif|webp);base64,)/i.test(value.trim()) ? value.trim() : '' }

function inlineRunsFromElement(root: HTMLElement, defaults: RunPatch = {}): DocumentRun[] {
  const runs: DocumentRun[] = []
  const walk = (node: Node, patch: RunPatch) => {
    if (node.nodeType === Node.TEXT_NODE) { if (node.textContent) runs.push(createRun(node.textContent, patch)); return }
    if (!(node instanceof HTMLElement)) return
    if (node.tagName === 'BR') { runs.push(createRun('\n', patch)); return }
    if (node.dataset.officeLatex === 'v1' && node.dataset.officeLatexMode !== 'block') {
      const latex = node.dataset.officeLatexSource ?? node.textContent ?? ''
      const run = createRun('\uFFFC', { ...patch, fontFamily: 'Cambria Math' })
      run.metadata = { ...run.metadata, latexFormula: { latex, mode: 'inline', version: 1 } }
      runs.push(run)
      return
    }
    const style = node.style
    const fontSize = Number.parseFloat(style.fontSize)
    const next: RunPatch = {
      ...patch,
      bold: patch.bold || ['B', 'STRONG'].includes(node.tagName) || Number(style.fontWeight) >= 600 || style.fontWeight === 'bold',
      italic: patch.italic || ['I', 'EM'].includes(node.tagName) || style.fontStyle === 'italic',
      underline: patch.underline || node.tagName === 'U' || style.textDecoration.includes('underline'),
      strike: patch.strike || ['S', 'STRIKE'].includes(node.tagName) || style.textDecoration.includes('line-through'),
      verticalAlign: node.tagName === 'SUP' ? 'superscript' : node.tagName === 'SUB' ? 'subscript' : normalizeVerticalAlign(style.verticalAlign) ?? patch.verticalAlign,
      color: style.color || node.getAttribute('color') || patch.color,
      fontFamily: style.fontFamily || node.getAttribute('face') || patch.fontFamily,
      fontSize: fontSize ? (style.fontSize.endsWith('px') ? fontSize * 0.75 : fontSize) : patch.fontSize,
      highlight: style.backgroundColor || patch.highlight,
      href: node.tagName === 'A' ? safeUrl(node.getAttribute('href') ?? '') : patch.href,
    }
    node.childNodes.forEach((child) => walk(child, next))
  }
  const { id: _id, text: _text, ...defaultMarks } = createRun('', defaults)
  walk(root, defaultMarks)
  return compactRuns(runs)
}

function normalizeVerticalAlign(value: string): DocumentRun['verticalAlign'] | undefined {
  if (value === 'super' || value === 'superscript') return 'superscript'
  if (value === 'sub' || value === 'subscript') return 'subscript'
  if (value === 'baseline') return 'baseline'
  return undefined
}

const escapeHtml = (text: string): string => text.replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char] ?? char)
function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
function runsHtml(runs: readonly InlineRun[]): string {
  return runs.map((source) => {
    const run = source as DocumentRun
    const formula = run.metadata?.latexFormula
    if (isRecord(formula) && typeof formula.latex === 'string' && (formula.mode === 'inline' || formula.mode === 'block')) {
      const latex = escapeHtml(formula.latex)
      return `<span data-office-latex="v1" data-office-latex-mode="${formula.mode}" data-office-latex-source="${latex}">${latex}</span>`
    }
    let value = escapeHtml(run.text).replace(/\n/g, '<br>')
    if (run.bold) value = `<strong>${value}</strong>`
    if (run.italic) value = `<em>${value}</em>`
    if (run.underline) value = `<u>${value}</u>`
    if (run.strike) value = `<s>${value}</s>`
    if (run.verticalAlign === 'superscript') value = `<sup>${value}</sup>`
    else if (run.verticalAlign === 'subscript') value = `<sub>${value}</sub>`
    value = `<span style="font-family:${escapeHtml(run.fontFamily)};font-size:${run.fontSize}pt;color:${escapeHtml(run.color)}${run.highlight ? `;background-color:${escapeHtml(run.highlight)}` : ''}">${value}</span>`
    return run.href ? `<a href="${escapeHtml(run.href)}">${value}</a>` : value
  }).join('') || '<br>'
}

export function documentHtml(model: DocumentModel): string {
  return model.blocks.map((block, index) => {
    const codeBlock = block.metadata.codeBlock
    if (isRecord(codeBlock) && typeof codeBlock.source === 'string') {
      const language = typeof codeBlock.language === 'string' ? codeBlock.language : 'plain'
      const source = escapeHtml(codeBlock.source)
      return `<pre data-office-code-block="v1" data-office-language="${escapeHtml(language)}" data-office-line-numbers="${codeBlock.lineNumbers === true}" style="white-space:pre-wrap;font-family:Consolas,monospace"><code>${source}</code></pre>`
    }
    const latexBlock = block.metadata.latexFormulaBlock === true ? (block.runs[0] as DocumentRun | undefined)?.metadata?.latexFormula : undefined
    if (isRecord(latexBlock) && typeof latexBlock.latex === 'string') {
      const source = escapeHtml(latexBlock.latex)
      return `<div data-office-latex="v1" data-office-latex-mode="block" data-office-latex-source="${source}" style="text-align:center">${source}</div>`
    }
    const table = tableForBlock(block)
    if (table) return `<table style="border-collapse:collapse;width:100%"><tbody>${table.rows.map((row) => `<tr>${row.cells.map((cell) => `<td rowspan="${cell.rowSpan}" colspan="${cell.colSpan}" style="border:1px solid #aab3bf;background:${escapeHtml(String(cell.style.fill ?? '#ffffff'))};padding:8px">${cell.blocks.map((paragraph) => `<p style="text-align:${paragraph.style.align}">${runsHtml(paragraph.runs)}</p>`).join('')}</td>`).join('')}</tr>`).join('')}</tbody></table>`
    const image = imageForBlock(block)
    if (image) return `<p><img src="${escapeHtml(image.src)}" alt="${escapeHtml(image.alt)}"${image.width ? ` width="${image.width}"` : ''}${image.height ? ` height="${image.height}"` : ''}></p>`
    if (block.type === 'pageBreak') return '<hr data-office-page-break="true" style="break-after:page">'
    if (block.type === 'columnBreak') return '<hr data-office-column-break="true">'
    if (block.type === 'sectionBreak') return `<hr data-office-section-break="${block.sectionBreak?.type ?? 'nextPage'}">`
    if (block.type === 'divider') return '<hr>'
    if (block.type === 'checklist') {
      const listId = block.checklist?.listId ?? block.id
      const previous = model.blocks[index - 1]
      if (previous?.type === 'checklist' && (previous.checklist?.listId ?? previous.id) === listId) return ''
      const items: string[] = []
      for (let cursor = index; cursor < model.blocks.length; cursor++) {
        const item = model.blocks[cursor]
        if (item?.type !== 'checklist' || (item.checklist?.listId ?? item.id) !== listId) break
        const content = `<p style="text-align:${item.style.align};padding-left:${item.style.indent * 24}px;line-height:${item.style.lineHeight};margin:${item.style.spaceBefore}px 0 ${item.style.spaceAfter}px">${runsHtml(item.runs)}</p>`
        items.push(`<li data-office-checked="${item.checklist?.checked === true}">${content}</li>`)
      }
      return `<ul data-office-list-type="checklist" data-office-list-id="${escapeHtml(listId)}">${items.join('')}</ul>`
    }
    const tag = block.type === 'heading' ? `h${Math.max(1, Math.min(6, block.level))}` : block.type === 'quote' ? 'blockquote' : 'p'
    const content = `<${tag} style="text-align:${block.style.align};padding-left:${block.style.indent * 24}px;line-height:${block.style.lineHeight};margin:${block.style.spaceBefore}px 0 ${block.style.spaceAfter}px">${runsHtml(block.runs)}</${tag}>`
    return block.type === 'bullet' || block.type === 'numbered' ? `<${block.type === 'bullet' ? 'ul' : 'ol'}><li>${content}</li></${block.type === 'bullet' ? 'ul' : 'ol'}>` : content
  }).join('')
}
