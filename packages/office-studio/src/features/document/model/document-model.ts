import { createDocumentDefaultStyle, type DocumentBlock, type DocumentModel, type DocumentParagraphStyle, type InlineRun } from '../../../domain/model'
import { blockHasNumbering, blockNumberingLevelIndex } from './document-numbering'
import { documentTableTextMatrix, documentTableValuesToTsv } from './document-table-text'
import { documentTableCellBlocks, documentTableCellContentContainers, documentTableContinuationEntries } from './document-table-continuations'
import { documentTableGrid } from './document-table-grid'
import { DOCUMENT_TABLE_BORDER_SIDES, DOCUMENT_TABLE_DEFAULT_MARGIN, documentTableCellBorders, documentTableCellFill, documentTableCellMargin, documentTableRowProperties } from './document-table-style'
import { documentRunHasNoteReference } from './document-note-reference'
import { documentBlockChildContainers, documentStructuredChildContainers, isDocumentTextBlock, isDocumentStructuredBlock } from './document-structural-blocks'

export interface DocumentRun extends InlineRun { highlight?: string; href?: string }
export interface TextTarget { blockId: string; row?: number; col?: number; cellBlockId?: string }
export interface DocumentPoint extends TextTarget {
  offset: number
  /** Derived view affinity for repeated table headers; never a document coordinate. */
  pageAffinity?: number
}
export interface DocumentTableViewSelection {
  tableId: string
  range: { startRow: number; endRow: number; startColumn: number; endColumn: number }
  kind: 'table' | 'row' | 'column' | 'cell' | 'range'
  pageAffinity?: number
}
export interface DocumentSelection {
  anchor: DocumentPoint
  focus: DocumentPoint
  /** Derived structural selection; legal text points remain input fallbacks. */
  tableSelection?: DocumentTableViewSelection
}
/** All table geometry in the domain is measured in CSS pixels. */
export interface DocumentTableCellMargin { top: number; right: number; bottom: number; left: number }
export interface DocumentTableBorder { width: number; color: string; style: 'solid' | 'dot' | 'dash' }
export type DocumentTableBorderSide = 'top' | 'right' | 'bottom' | 'left'
export type DocumentTableBorders = Partial<Record<DocumentTableBorderSide | 'insideHorizontal' | 'insideVertical', DocumentTableBorder>>
export interface DocumentTableContinuationContent {
  version: 1
  ownerCellId: string
  rowId: string
  colSpan: number
  blocks: DocumentBlock[]
  source?: Record<string, unknown>
}
export interface DocumentTableCellStyle extends Record<string, unknown> {
  fill?: string
  background?: string
  verticalAlign?: 'top' | 'center' | 'bottom'
  margin?: DocumentTableCellMargin
  borders?: Partial<Record<DocumentTableBorderSide, DocumentTableBorder>>
  /** This project's dropdown editor stores option identity separately from its visible text. */
  selectedDropdownOptionId?: string
  /** Independent OOXML continuation content remains an editable, separately written container. */
  nativeContinuationContent?: Record<string, DocumentTableContinuationContent>
}
export interface DocumentTableStyle extends Record<string, unknown> {
  fill?: string
  background?: string
  align?: 'left' | 'center' | 'right'
  indent?: number
  offsetY?: number
  width?: number
  layout?: 'auto' | 'fixed'
  cellMargin?: DocumentTableCellMargin
  borders?: DocumentTableBorders
}
export interface DocumentTableColumnTypeConfig {
  type: 'text' | 'number' | 'date' | 'checkbox' | 'dropdown'
  dropdownOptions?: Array<{ id: string; label: string; color?: string }>
}
export interface DocumentCell {
  id: string
  blocks: DocumentBlock[]
  style: DocumentTableCellStyle
  /** Logical 0-based anchor column; TextTarget.col remains the stored cell index. */
  columnIndex?: number
  rowSpan: number
  colSpan: number
}
export interface DocumentTableRow {
  id: string
  /** Only anchor cells are stored. A row covered by earlier rowSpan cells can be empty. */
  cells: DocumentCell[]
  height?: number
  heightRule?: 'auto' | 'atLeast' | 'exact'
  allowSplit?: boolean
  style?: Record<string, unknown>
}
export interface DocumentTable {
  rows: DocumentTableRow[]
  columnWidths?: number[]
  headerRowCount?: number
  titleRow?: boolean
  columnTypes?: Record<number, DocumentTableColumnTypeConfig>
  style?: DocumentTableStyle
  /** Imported legacy row properties remain available to the layout/IO adapters. */
  rowStyles?: Array<Record<string, unknown>>
}
export interface DocumentImage { src: string; alt: string; width: number; height: number }
export type RunPatch = Partial<Omit<DocumentRun, 'id' | 'text'>>
export type DocumentRunClearField = 'color' | 'highlight'
export const DOCUMENT_AUTOMATIC_TEXT_COLOR = 'auto'
export const documentTextColor = (color: string): string => color === DOCUMENT_AUTOMATIC_TEXT_COLOR ? '#1f2937' : color

/** The command keeps its clear intent in JSON; undefined is only an in-memory run patch. */
export function documentRunFormatPatch(patch: RunPatch, clearFields: readonly DocumentRunClearField[] = []): RunPatch {
  const result = { ...patch }
  for (const field of clearFields) {
    if (field !== 'color' && field !== 'highlight') throw new RangeError('Unknown document text style field to clear')
    result[field] = undefined
  }
  return result
}

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
  const numberingLevel = blockHasNumbering(block) ? blockNumberingLevelIndex(block) : undefined
  block.metadata = { ...block.metadata, namedStyleType: namedStyle }
  if (numberingLevel !== undefined && block.type !== 'checklist') block.metadata.numberingLevel = numberingLevel
  block.type = style.level > 0 ? 'heading' : 'paragraph'
  block.level = style.level
  block.style = {
    ...block.style,
    indent: 0,
    lineHeight: 1.5,
    spacingRule: 'auto',
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
  const updated: DocumentRun = { ...run, ...patch, color: patch.color ?? run.color }
  for (const field of ['color', 'highlight'] as const) {
    if (!Object.hasOwn(patch, field) || patch[field] !== undefined) continue
    if (field === 'color') updated.color = DOCUMENT_AUTOMATIC_TEXT_COLOR
    else delete updated.highlight
    if (updated.metadata?.[field] !== undefined) {
      const metadata = { ...updated.metadata }
      delete metadata[field]
      if (Object.keys(metadata).length) updated.metadata = metadata
      else delete updated.metadata
    }
  }
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
export const createRun = (text = '', patch: RunPatch = {}): DocumentRun => applyRunPatch({ id: documentId('run'), text, ...createDocumentDefaultStyle().text }, patch)
export const createParagraph = (text = '', patch: RunPatch = {}): DocumentBlock => ({ id: documentId('paragraph'), type: 'paragraph', level: 0, runs: [createRun(text, patch)], style: defaultParagraphStyle(), metadata: {} })
export const createChecklistParagraph = (text = '', checked = false, listId = documentId('checklist')): DocumentBlock => ({ ...createParagraph(text), type: 'checklist', checklist: { listId, checked } })
export const targetKey = (target: TextTarget): string => target.row == null ? target.blockId : `${target.blockId}/r${target.row}/c${target.col ?? 0}/${target.cellBlockId ?? ''}`
export const sameTarget = (a: TextTarget, b: TextTarget): boolean => targetKey(a) === targetKey(b)
export const collapsedSelection = (point: DocumentPoint): DocumentSelection => ({ anchor: { ...point }, focus: { ...point } })
export const selectionCollapsed = (selection: DocumentSelection): boolean => !selection.tableSelection && sameTarget(selection.anchor, selection.focus) && selection.anchor.offset === selection.focus.offset
export const runText = (runs: readonly InlineRun[]): string => runs.map((run) => run.text).join('')

export function tableForBlock(block: DocumentBlock): DocumentTable | undefined {
  const table = block.metadata.table
  if (!table || typeof table !== 'object' || !('rows' in table) || !Array.isArray(table.rows)) return undefined
  return table as DocumentTable
}

/** Stable block IDs address paragraphs and tables at every nesting depth. */
export function findDocumentBlock(model: Pick<DocumentModel, 'blocks'>, id: string): DocumentBlock | undefined {
  const container = findDocumentBlockContainer(model, id)
  return container?.find(block => block.id === id)
}

export function findDocumentBlockContainer(model: Pick<DocumentModel, 'blocks'>, id: string): DocumentBlock[] | undefined {
  const visit = (blocks: DocumentBlock[]): DocumentBlock[] | undefined => {
    if (blocks.some(block => block.id === id)) return blocks
    for (const block of blocks) for (const content of documentBlockChildContainers(block)) {
      const nested = visit(content)
      if (nested) return nested
    }
  }
  return visit(model.blocks)
}

export function documentBlockAncestry(model: Pick<DocumentModel, 'blocks'>, id: string): DocumentBlock[] | undefined {
  const visit = (blocks: DocumentBlock[]): DocumentBlock[] | undefined => {
    for (const block of blocks) {
      if (block.id === id) return [block]
      for (const content of documentBlockChildContainers(block)) {
        const nested = visit(content)
        if (nested) return [block, ...nested]
      }
    }
  }
  return visit(model.blocks)
}

export function imageForBlock(block: DocumentBlock): DocumentImage | undefined {
  const image = block.metadata.image
  if (!image || typeof image !== 'object' || !('src' in image) || typeof image.src !== 'string') return undefined
  return image as DocumentImage
}

export function textTargets(model: DocumentModel): TextTarget[] {
  const targets: TextTarget[] = []
  const visit = (blocks: DocumentBlock[], host?: Omit<TextTarget, 'cellBlockId'>) => {
    for (const block of blocks) {
      const table = tableForBlock(block)
      if (block.type === 'table' && table) {
        table.rows.forEach((row, rowIndex) => row.cells.forEach((cell, cellIndex) => {
          for (const content of documentTableCellContentContainers(cell)) visit(content, { blockId: block.id, row: rowIndex, col: cellIndex })
        }))
      } else if (isDocumentStructuredBlock(block)) {
        for (const content of documentStructuredChildContainers(block)) visit(content, host)
      } else if (isDocumentTextBlock(block)) {
        targets.push(host ? { ...host, cellBlockId: block.id } : { blockId: block.id })
      }
    }
  }
  visit(model.blocks)
  return targets
}

export function blockAtPoint(model: DocumentModel, point: TextTarget): DocumentBlock | undefined {
  const block = findDocumentBlock(model, point.blockId)
  if (point.row == null || !block) return block
  const cell = tableForBlock(block)?.rows[point.row]?.cells[point.col ?? 0]
  if (!cell) return undefined
  const contents = documentTableCellBlocks(cell)
  if (point.cellBlockId) {
    const candidate = findDocumentBlock({ blocks: contents }, point.cellBlockId)
    if (candidate && isDocumentTextBlock(candidate)) return candidate
  }
  const visit = (blocks: DocumentBlock[]): DocumentBlock | undefined => {
    for (const candidate of blocks) {
      if (isDocumentTextBlock(candidate)) return candidate
      for (const children of documentStructuredChildContainers(candidate)) {
        const found = visit(children)
        if (found) return found
      }
    }
  }
  return visit(contents)
}

export function runsForTarget(model: DocumentModel, target: TextTarget): DocumentRun[] {
  return (blockAtPoint(model, target)?.runs ?? []) as DocumentRun[]
}

export function setTargetRuns(model: DocumentModel, target: TextTarget, runs: DocumentRun[]): void {
  const block = blockAtPoint(model, target)
  if (!block) return
  block.runs = compactRuns(runs)
  delete block.metadata.html
  for (const ancestor of documentBlockAncestry(model, block.id) ?? []) delete ancestor.metadata.html
}

export function compactRuns(runs: DocumentRun[]): DocumentRun[] {
  const result: DocumentRun[] = []
  for (const source of runs) {
    const entity = documentRunHasNoteReference(source) || source.metadata?.latexFormula !== undefined
    if (!source.text && !entity) continue
    const normalized = source.verticalAlign === 'baseline' ? applyRunPatch(source, { verticalAlign: 'baseline' }) : source
    const last = result.at(-1)
    const { id: _id, text: _text, ...marks } = normalized
    if (last) {
      const { id: _lastId, text: _lastText, ...lastMarks } = last
      if (!entity && !documentRunHasNoteReference(last) && last.metadata?.latexFormula === undefined
        && JSON.stringify(marks) === JSON.stringify(lastMarks)) { last.text += normalized.text; continue }
    }
    result.push({ ...normalized })
  }
  return result.length ? result : [createRun()]
}

export function sliceRuns(runs: readonly DocumentRun[], start: number, end = Number.POSITIVE_INFINITY): DocumentRun[] {
  const result: DocumentRun[] = []
  let offset = 0
  for (const run of runs) {
    // Half-open ownership includes a zero-width reference in its suffix once.
    if (!run.text && (documentRunHasNoteReference(run) || run.metadata?.latexFormula !== undefined)) {
      if (offset >= start && offset < end) result.push({ ...run })
      continue
    }
    const from = Math.max(0, start - offset)
    const to = Math.min(run.text.length, end - offset)
    if (to > from) result.push({ ...run, id: from === 0 && to === run.text.length ? run.id : documentId('run'), text: run.text.slice(from, to) })
    offset += run.text.length
    if (offset > end) break
  }
  return result
}

export function clampPoint(model: DocumentModel, point: DocumentPoint): DocumentPoint {
  const targets = textTargets(model)
  const block = point.row == null ? undefined : findDocumentBlock(model, point.blockId)
  const table = block && tableForBlock(block)
  const row = table ? Math.max(0, Math.min(point.row ?? 0, table.rows.length - 1)) : undefined
  const col = row == null ? undefined : Math.max(0, Math.min(point.col ?? 0, (table?.rows[row]?.cells.length ?? 1) - 1))
  // Undo/redo can replace a cell paragraph id while retaining its logical table coordinate.
  const target = targets.find((candidate) => sameTarget(candidate, point))
    ?? (point.cellBlockId ? targets.find(candidate => candidate.cellBlockId === point.cellBlockId) : undefined)
    ?? targets.find((candidate) => candidate.blockId === point.blockId && candidate.row === row && candidate.col === col)
    ?? targets.find((candidate) => candidate.blockId === point.blockId)
    ?? targets[0]
  return target ? { ...target, offset: Math.max(0, Math.min(runText(runsForTarget(model, target)).length, point.offset)),
    ...(Number.isInteger(point.pageAffinity) && point.pageAffinity! >= 0 && sameTarget(target, point) ? { pageAffinity: point.pageAffinity } : {}) }
    : { blockId: model.blocks[0]?.id ?? '', offset: 0 }
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
  if (patch.metadata) {
    const metadata = { ...patch.metadata }
    for (const key of ['note', 'noteReference', 'latexFormula', 'officeStoryRunSourceId']) delete metadata[key]
    if (Object.keys(metadata).length) patch.metadata = metadata
    else delete patch.metadata
  }
  return patch
}

export function selectedPlainText(model: DocumentModel, selection: DocumentSelection): string {
  if (selection.tableSelection) {
    const block = findDocumentBlock(model, selection.tableSelection.tableId), table = block && tableForBlock(block)
    if (!table) throw new RangeError('表格结构选区已不存在')
    return documentTableValuesToTsv(documentTableTextMatrix(table, selection.tableSelection.range))
  }
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
  const normalizeBlocks = (blocks: DocumentBlock[]): DocumentBlock[] => blocks.flatMap((block) => {
    if (isDocumentStructuredBlock(block)) {
      if (block.children) block.children = normalizeBlocks(block.children)
      if (block.columns) for (const column of block.columns.columns) column.blocks = normalizeBlocks(column.blocks)
      delete block.metadata.html
      return [block]
    }
    const html = block.metadata.html
    if (typeof html !== 'string') return [block]
    const parsed = blocksFromHtml(html)
    if (!parsed.length) return [block]
    const first = parsed[0]
    return [{ ...first, id: block.id, type: block.type === 'paragraph' ? first.type : block.type, level: block.level || first.level, style: { ...first.style, ...block.style }, metadata: { ...block.metadata, ...first.metadata, sourceHtml: html, html: undefined } }, ...parsed.slice(1)]
  })
  result.blocks = normalizeBlocks(result.blocks)
  const normalizeProperties = (blocks: DocumentBlock[]) => { for (const block of blocks) {
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
    for (const container of documentBlockChildContainers(block)) normalizeProperties(container)
  } }
  normalizeProperties(result.blocks)
  if (!textTargets(result).length) result.blocks.push(createParagraph())
  result.wordCount = textTargets(result).reduce((count, target) => count + runText(runsForTarget(result, target)).replace(/\s/g, '').length, 0)
  return result
}

export function createTable(rows: number, cols: number): DocumentTable {
  return { rows: Array.from({ length: rows }, () => ({ id: documentId('row'), cells: Array.from({ length: cols }, (_, columnIndex) => ({ id: documentId('cell'), blocks: [createParagraph()], style: {}, columnIndex, rowSpan: 1, colSpan: 1 })) })), columnWidths: [], headerRowCount: 0 }
}

/** HTML is accepted at import/clipboard boundaries only; it is never the editor state. */
export function blocksFromHtml(html: string): DocumentBlock[] {
  const root = document.createElement('div')
  root.innerHTML = html
  root.querySelectorAll('script,style,iframe,object,embed,meta,link,title,base').forEach((node) => node.remove())
  let blockCount = 0, logicalCells = 0
  const structural = 'p,div,h1,h2,h3,h4,h5,h6,blockquote,ul,ol,pre,table,hr,img,section,article,main'
  const countBlock = (block: DocumentBlock): DocumentBlock => {
    if (++blockCount > 250_000) throw new RangeError('HTML 粘贴超过单次文档块预算')
    return block
  }
  const parseChildren = (container: HTMLElement, defaults: RunPatch, paragraphStyle: DocumentParagraphStyle, depth: number,
    listDepth = 0, excluded: ReadonlySet<Node> = new Set()): DocumentBlock[] => {
    if (depth > 64) throw new RangeError('HTML 粘贴超过文档嵌套预算')
    const blocks: DocumentBlock[] = [], pending: Node[] = []
    const flush = () => {
      if (!pending.length) return
      const wrapper = container.cloneNode(false) as HTMLElement
      for (const node of pending) wrapper.append(node.cloneNode(true))
      pending.length = 0
      const runs = inlineRunsFromElement(wrapper, defaults)
      if (!runText(runs).trim() && !runs.some(run => /[\n\uFFFC]/.test(run.text))) return
      blocks.push(countBlock({ ...createParagraph(), runs, style: { ...paragraphStyle } }))
    }
    for (const node of Array.from(container.childNodes)) {
      if (excluded.has(node)) continue
      if (node.nodeType === Node.TEXT_NODE) { pending.push(node); continue }
      if (!(node instanceof HTMLElement) || ['INPUT', 'BUTTON', 'SELECT', 'TEXTAREA'].includes(node.tagName)) continue
      if (node.matches(structural) || node.querySelector(structural)) {
        flush()
        blocks.push(...append(node, defaults, paragraphStyle, depth + 1, listDepth))
      } else pending.push(node)
    }
    flush()
    return blocks
  }
  const append = (element: HTMLElement, defaults: RunPatch, inheritedStyle: DocumentParagraphStyle, depth: number, listDepth: number): DocumentBlock[] => {
    if (depth > 64) throw new RangeError('HTML 粘贴超过文档嵌套预算')
    const tag = element.tagName
    const marks = htmlElementRunPatch(element, defaults), paragraphStyle = htmlParagraphStyle(element, inheritedStyle)
    if (tag === 'PRE' && element.dataset.officeCodeBlock === 'v1') {
      const source = element.querySelector('code')?.textContent ?? element.textContent ?? ''
      const languageValue = element.dataset.officeLanguage ?? 'plain'
      const language = ['plain', 'javascript', 'typescript', 'json', 'css', 'html', 'python', 'sql', 'rust'].includes(languageValue) ? languageValue : 'plain'
      const block = createParagraph(source, { fontFamily: 'Consolas', fontSize: 10 })
      block.metadata.codeBlock = { source, language, lineNumbers: element.dataset.officeLineNumbers === 'true' }
      block.style = { ...block.style, align: 'left', indent: 0, lineHeight: 1.35, spaceBefore: 8, spaceAfter: 8 }
      return [countBlock(block)]
    }
    if (tag === 'DIV' && element.dataset.officeLatex === 'v1' && element.dataset.officeLatexMode === 'block') {
      const latex = element.dataset.officeLatexSource ?? element.textContent ?? ''
      const run = createRun('\uFFFC', { fontFamily: 'Cambria Math' })
      run.metadata = { ...run.metadata, latexFormula: { latex, mode: 'block', version: 1 } }
      const block = createParagraph()
      block.runs = [run]
      block.style = { ...block.style, align: 'center', spaceBefore: 8, spaceAfter: 8 }
      block.metadata.latexFormulaBlock = true
      return [countBlock(block)]
    }
    if (tag === 'UL' || tag === 'OL') {
      const blocks: DocumentBlock[] = []
      const checklist = tag === 'UL' && element.dataset.officeListType === 'checklist'
      const listId = checklist ? element.dataset.officeListId || documentId('checklist') : undefined
      for (const child of Array.from(element.children)) {
        if (!(child instanceof HTMLElement) || child.tagName !== 'LI') continue
        const item = child, itemStyle = { ...paragraphStyle, indent: paragraphStyle.indent + (listDepth > 0 ? 1 : 0) }
        const checkbox = Array.from(item.querySelectorAll<HTMLInputElement>('input[type="checkbox"]')).find(input => input.closest('li') === item)
        const isChecklistItem = checklist || checkbox !== undefined
        const children = parseChildren(item, htmlElementRunPatch(item, marks), itemStyle, depth + 1, listDepth + 1)
        const block = children.find(candidate => !['table', 'image', 'divider', 'pageBreak', 'columnBreak', 'sectionBreak', 'bullet', 'numbered', 'checklist'].includes(candidate.type))
          ?? countBlock(createParagraph())
        if (!children.includes(block)) children.unshift(block)
        block.type = isChecklistItem ? 'checklist' : tag === 'UL' ? 'bullet' : 'numbered'
        if (isChecklistItem) block.checklist = {
          listId: listId ?? documentId('checklist'),
          checked: item.dataset.officeChecked === 'true' || checkbox?.checked === true || checkbox?.hasAttribute('checked') === true,
        }
        block.style = { ...itemStyle, ...block.style, indent: itemStyle.indent }
        blocks.push(...children)
      }
      return blocks
    }
    if (tag === 'TABLE') return [countBlock(parseTable(element as HTMLTableElement, marks, paragraphStyle, depth + 1))]
    if (tag === 'IMG') {
      const image = element as HTMLImageElement, block = createParagraph()
      block.type = 'image'; block.runs = []
      block.metadata.image = { src: safeUrl(image.getAttribute('src') ?? ''), alt: image.alt,
        width: htmlCssPixels(image.style.width || image.getAttribute('width')) ?? 0,
        height: htmlCssPixels(image.style.height || image.getAttribute('height')) ?? 0 } satisfies DocumentImage
      return [countBlock(block)]
    }
    if (element.querySelector(structural)) {
      const children = parseChildren(element, marks, paragraphStyle, depth + 1, listDepth)
      if (tag === 'BLOCKQUOTE') for (const block of children) if (block.type === 'paragraph') block.type = 'quote'
      if (/^H[1-6]$/.test(tag)) for (const block of children) if (block.type === 'paragraph') { block.type = 'heading'; block.level = Number(tag.slice(1)) }
      return children
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
    block.runs = inlineRunsFromElement(element, block.type === 'heading' ? { ...defaults, fontSize: block.level === 1 ? 24 : block.level === 2 ? 18 : 14, bold: true } : defaults)
    block.style = paragraphStyle
    return [countBlock(block)]
  }
  const parseTable = (element: HTMLTableElement, defaults: RunPatch, paragraphStyle: DocumentParagraphStyle, depth: number): DocumentBlock => {
    const sourceRows = Array.from(element.rows).filter(row => row.closest('table') === element)
    if (!sourceRows.length || sourceRows.length > 10_000) throw new RangeError('HTML 表格行数必须在 1 到 10000 之间')
    const table: DocumentTable = { rows: sourceRows.map(() => ({ id: documentId('row'), cells: [] })), columnWidths: [] }
    const occupied: Array<Array<{ start: number; end: number }>> = sourceRows.map(() => [])
    const measuredWidths: Array<number | undefined> = []
    const groupEnd = new Map<Element | null, number>()
    sourceRows.forEach((row, index) => groupEnd.set(row.parentElement, index))
    let columnCount = 0, generatedCells = 0
    for (const [rowIndex, row] of sourceRows.entries()) {
      const height = htmlCssPixels(row.style.height || row.getAttribute('height')), rule = row.dataset.officeRowHeightRule
      if (height !== undefined && height > 0) table.rows[rowIndex]!.height = height
      table.rows[rowIndex]!.heightRule = rule === 'auto' || rule === 'atLeast' || rule === 'exact' ? rule : height ? 'atLeast' : 'auto'
      table.rows[rowIndex]!.allowSplit = row.dataset.officeRowAllowSplit !== 'false' && row.style.breakInside !== 'avoid' && row.style.pageBreakInside !== 'avoid'
      let cursor = 0
      for (const elementCell of Array.from(row.cells)) {
        const rowSpan = elementCell.rowSpan === 0 ? groupEnd.get(row.parentElement)! - rowIndex + 1 : elementCell.rowSpan
        const colSpan = elementCell.colSpan
        if (!Number.isSafeInteger(rowSpan) || rowSpan < 1 || rowIndex + rowSpan > sourceRows.length) throw new RangeError('HTML 表格跨行跨度超出实际行数')
        occupied[rowIndex]!.sort((left, right) => left.start - right.start)
        for (const interval of occupied[rowIndex]!) {
          if (interval.end <= cursor) continue
          if (cursor + colSpan <= interval.start) break
          cursor = interval.end
        }
        if (cursor + colSpan > 512 || sourceRows.length * (cursor + colSpan) > 1_000_000) throw new RangeError('HTML 表格超出逻辑网格预算')
        const cell: DocumentCell = { id: documentId('cell'), columnIndex: cursor, rowSpan, colSpan, style: htmlTableCellStyle(elementCell), blocks: [] }
        const continuationElements = Array.from(elementCell.children).filter((child): child is HTMLElement => child instanceof HTMLElement && child.dataset.officeTableContinuation === 'v1')
        const excluded = new Set<Node>(continuationElements), cellMarks = htmlElementRunPatch(elementCell, defaults)
        const cellParagraphStyle = htmlParagraphStyle(elementCell, paragraphStyle, false)
        cell.blocks = parseChildren(elementCell, cellMarks, cellParagraphStyle, depth + 1, 0, excluded)
        if (!cell.blocks.length) cell.blocks.push(countBlock({ ...createParagraph('', cellMarks), style: { ...cellParagraphStyle } }))
        const continuations: Record<string, DocumentTableContinuationContent> = {}
        for (const continuation of continuationElements) {
          const key = continuation.dataset.officeRowOffset ?? '', offset = Number(key)
          if (!Number.isSafeInteger(offset) || offset < 1 || String(offset) !== key || offset >= rowSpan || continuations[key]) throw new RangeError('HTML 表格续格偏移无效')
          continuations[key] = { version: 1, ownerCellId: cell.id, rowId: table.rows[rowIndex + offset]!.id, colSpan,
            blocks: parseChildren(continuation, cellMarks, cellParagraphStyle, depth + 1) }
        }
        if (continuationElements.length) cell.style.nativeContinuationContent = continuations
        table.rows[rowIndex]!.cells.push(cell)
        const width = htmlCssPixels(elementCell.style.width || elementCell.getAttribute('width'))
        if (colSpan === 1 && width && width > 0) measuredWidths[cursor] ??= width
        for (let coveredRow = rowIndex; coveredRow < rowIndex + rowSpan; coveredRow++) {
          if (occupied[coveredRow]!.some(interval => interval.start < cursor + colSpan && cursor < interval.end)) throw new RangeError('HTML 表格合并跨度重叠')
          occupied[coveredRow]!.push({ start: cursor, end: cursor + colSpan })
        }
        cursor += colSpan; columnCount = Math.max(columnCount, cursor)
      }
    }
    let colIndex = 0
    for (const col of Array.from(element.querySelectorAll('col')).filter(col => col.closest('table') === element)) {
      const span = col.span, width = htmlCssPixels(col.style.width || col.getAttribute('width'))
      if (colIndex + span > 512) throw new RangeError('HTML 表格列声明超出预算')
      for (let offset = 0; offset < span; offset++) if (width && width > 0) measuredWidths[colIndex + offset] = width
      colIndex += span
    }
    columnCount = Math.max(columnCount, colIndex)
    if (!columnCount) throw new RangeError('HTML 表格至少需要一个单元格')
    logicalCells += sourceRows.length * columnCount
    if (logicalCells > 250_000) throw new RangeError('HTML 粘贴超过单次单元格预算')
    for (const [rowIndex, intervals] of occupied.entries()) {
      intervals.sort((left, right) => left.start - right.start)
      let column = 0
      for (const interval of [...intervals, { start: columnCount, end: columnCount }]) {
        for (; column < interval.start; column++) {
          table.rows[rowIndex]!.cells.push({ id: documentId('cell'), columnIndex: column, rowSpan: 1, colSpan: 1, style: {}, blocks: [countBlock(createParagraph())] })
          generatedCells++
        }
        column = interval.end
      }
      table.rows[rowIndex]!.cells.sort((left, right) => left.columnIndex! - right.columnIndex!)
    }
    const width = htmlCssPixels(element.style.width || element.getAttribute('width')), headerCount = element.dataset.officeHeaderRowCount
    table.headerRowCount = headerCount === undefined ? sourceRows.filter(row => row.parentElement?.tagName === 'THEAD').length : Number(headerCount)
    const align = element.dataset.officeTableAlign || element.getAttribute('align') || element.style.textAlign
    const indent = element.dataset.officeTableIndent === undefined ? htmlCssPixels(element.style.marginLeft) : Number(element.dataset.officeTableIndent)
    const offsetY = element.dataset.officeTableOffsetY === undefined ? undefined : Number(element.dataset.officeTableOffsetY)
    table.style = { layout: element.style.tableLayout === 'fixed' ? 'fixed' : 'auto',
      ...(width && width > 0 ? { width } : {}), ...(align === 'left' || align === 'center' || align === 'right' ? { align } : {}),
      ...(indent !== undefined && Number.isFinite(indent) ? { indent } : {}), ...(offsetY !== undefined && Number.isFinite(offsetY) ? { offsetY } : {}),
      ...(element.style.backgroundColor ? { fill: element.style.backgroundColor } : {}) }
    if (measuredWidths.some(value => value !== undefined)) {
      const widths = Array.from({ length: columnCount }, (_, column) => measuredWidths[column])
      const total = widths.reduce<number>((sum, value) => sum + (value ?? 0), 0), missing = widths.filter(value => value === undefined).length
      const fallback = width && width > total && missing ? (width - total) / missing : total / Math.max(1, columnCount - missing)
      table.columnWidths = widths.map(value => value ?? fallback)
    }
    const block = createParagraph(); block.type = 'table'; block.runs = []
    block.metadata = { table, sourceHtml: element.outerHTML, ...(generatedCells ? { htmlTableNormalization: { generatedCells } } : {}) }
    documentTableGrid(table)
    return block
  }
  return parseChildren(root, {}, defaultParagraphStyle(), 0)
}

function htmlCssPixels(value: string | null | undefined): number | undefined {
  const match = value?.trim().match(/^([+-]?(?:\d+\.?\d*|\.\d+))(px|pt|in|cm|mm)?$/i)
  if (!match) return undefined
  const number = Number(match[1]), unit = match[2]?.toLowerCase()
  const scale = unit === 'pt' ? 4 / 3 : unit === 'in' ? 96 : unit === 'cm' ? 96 / 2.54 : unit === 'mm' ? 96 / 25.4 : 1
  return Number.isFinite(number) ? number * scale : undefined
}

function htmlParagraphStyle(element: HTMLElement, inherited: DocumentParagraphStyle, paragraphIndent = true): DocumentParagraphStyle {
  const style = { ...inherited }, css = element.style, lineHeight = Number(css.lineHeight)
  if (css.textAlign) style.align = textAlign(css.textAlign)
  const indent = htmlCssPixels(css.paddingLeft), before = htmlCssPixels(css.marginTop), after = htmlCssPixels(css.marginBottom)
  if (paragraphIndent && indent !== undefined) style.indent = indent / 24
  if (lineHeight > 0 && Number.isFinite(lineHeight)) style.lineHeight = lineHeight
  if (before !== undefined) style.spaceBefore = before
  if (after !== undefined) style.spaceAfter = after
  return style
}

function htmlTableCellStyle(element: HTMLTableCellElement): DocumentTableCellStyle {
  const css = element.style, fill = css.backgroundColor || element.getAttribute('bgcolor'), style: DocumentTableCellStyle = {}
  if (fill) style.fill = fill
  if (css.verticalAlign === 'top' || css.verticalAlign === 'bottom') style.verticalAlign = css.verticalAlign
  else if (css.verticalAlign === 'middle' || css.verticalAlign === 'center') style.verticalAlign = 'center'
  const margin = { ...DOCUMENT_TABLE_DEFAULT_MARGIN }
  let hasMargin = false
  for (const side of DOCUMENT_TABLE_BORDER_SIDES) {
    const value = htmlCssPixels(css.getPropertyValue('padding-' + side))
    if (value !== undefined && value >= 0) { margin[side] = value; hasMargin = true }
    const kind = css.getPropertyValue('border-' + side + '-style'), width = htmlCssPixels(css.getPropertyValue('border-' + side + '-width'))
    const color = css.getPropertyValue('border-' + side + '-color')
    if (kind || width !== undefined || color) {
      style.borders ??= {}
      style.borders[side] = { width: kind === 'none' || kind === 'hidden' ? 0 : width ?? 1,
        color: color || '#aab3bf', style: kind === 'dotted' ? 'dot' : kind === 'dashed' ? 'dash' : 'solid' }
    }
  }
  if (hasMargin) style.margin = margin
  return style
}

function textAlign(value: string): DocumentParagraphStyle['align'] { return value === 'center' || value === 'right' || value === 'justify' ? value : 'left' }
export function safeUrl(value: string): string { return /^(https?:|mailto:|#|blob:|data:image\/(png|jpeg|gif|webp);base64,)/i.test(value.trim()) ? value.trim() : '' }

function htmlElementRunPatch(node: HTMLElement, inherited: RunPatch): RunPatch {
  const next = { ...inherited }, style = node.style, decoration = style.textDecorationLine || style.textDecoration
  if (['B', 'STRONG'].includes(node.tagName)) next.bold = true
  else if (style.fontWeight) next.bold = Number(style.fontWeight) >= 600 || style.fontWeight === 'bold' || style.fontWeight === 'bolder'
  if (['I', 'EM'].includes(node.tagName)) next.italic = true
  else if (style.fontStyle) next.italic = style.fontStyle === 'italic' || style.fontStyle === 'oblique'
  if (node.tagName === 'U') next.underline = true
  else if (decoration) next.underline = decoration.includes('underline')
  if (['S', 'STRIKE'].includes(node.tagName)) next.strike = true
  else if (decoration) next.strike = decoration.includes('line-through')
  const vertical = node.tagName === 'SUP' ? 'superscript' : node.tagName === 'SUB' ? 'subscript' : normalizeVerticalAlign(style.verticalAlign)
  if (vertical) next.verticalAlign = vertical
  const color = style.color || node.getAttribute('color')
  if (node.dataset.officeTextColor === 'auto') next.color = DOCUMENT_AUTOMATIC_TEXT_COLOR
  else if (color && color !== 'inherit') next.color = color
  const family = style.fontFamily || node.getAttribute('face'), size = htmlCssPixels(style.fontSize)
  if (family && family !== 'inherit') next.fontFamily = family
  if (size && size > 0) next.fontSize = size * 0.75
  if (!['TABLE', 'TR', 'TD', 'TH', 'THEAD', 'TBODY', 'TFOOT'].includes(node.tagName)
    && style.backgroundColor && style.backgroundColor !== 'inherit') next.highlight = style.backgroundColor
  if (node.tagName === 'A') next.href = safeUrl(node.getAttribute('href') ?? '')
  return next
}

function inlineRunsFromElement(root: HTMLElement, defaults: RunPatch = {}): DocumentRun[] {
  const runs: DocumentRun[] = []
  const walk = (node: Node, patch: RunPatch) => {
    if (node.nodeType === Node.TEXT_NODE) { if (node.textContent) runs.push(createRun(node.textContent, patch)); return }
    if (!(node instanceof HTMLElement) || ['INPUT', 'BUTTON', 'SELECT', 'TEXTAREA'].includes(node.tagName)) return
    if (node.tagName === 'BR') { runs.push(createRun('\n', patch)); return }
    if (node.dataset.officeLatex === 'v1' && node.dataset.officeLatexMode !== 'block') {
      const latex = node.dataset.officeLatexSource ?? node.textContent ?? ''
      const run = createRun('\uFFFC', { ...patch, fontFamily: 'Cambria Math' })
      run.metadata = { ...run.metadata, latexFormula: { latex, mode: 'inline', version: 1 } }
      runs.push(run)
      return
    }
    const next = htmlElementRunPatch(node, patch)
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
    const automaticColor = run.color === DOCUMENT_AUTOMATIC_TEXT_COLOR
    value = `<span${automaticColor ? ' data-office-text-color="auto"' : ''} style="font-family:${escapeHtml(run.fontFamily)};font-size:${run.fontSize}pt;color:${automaticColor ? 'inherit' : escapeHtml(run.color)}${run.highlight ? `;background-color:${escapeHtml(run.highlight)}` : ''}">${value}</span>`
    return run.href ? `<a href="${escapeHtml(run.href)}">${value}</a>` : value
  }).join('') || '<br>'
}

export function documentHtml(model: DocumentModel): string {
  return documentBlocksHtml(model.blocks)
}

function documentTableHtml(table: DocumentTable): string {
  const grid = documentTableGrid(table), align = table.style?.align ?? 'left'
  const tableStyles = ['border-collapse:collapse', 'table-layout:' + (table.style?.layout ?? 'fixed'),
    'width:' + (table.style?.width ? table.style.width + 'px' : '100%')]
  if (align === 'center') tableStyles.push('margin-left:auto', 'margin-right:auto')
  else if (align === 'right') tableStyles.push('margin-left:auto', 'margin-right:0')
  else if (table.style?.indent) tableStyles.push('margin-left:' + table.style.indent + 'px')
  const attributes = ' data-office-header-row-count="' + (table.headerRowCount ?? 0) + '" data-office-table-align="' + align
    + '" data-office-table-indent="' + (table.style?.indent ?? 0) + '" data-office-table-offset-y="' + (table.style?.offsetY ?? 0) + '"'
  const columns = table.columnWidths?.length ? '<colgroup>' + table.columnWidths.map(width => '<col style="width:' + width + 'px">').join('') + '</colgroup>' : ''
  const rows = table.rows.map((row, rowIndex) => {
    const properties = documentTableRowProperties(table, rowIndex)
    const rowStyles = [properties.height ? 'height:' + properties.height + 'px' : '', !properties.allowSplit ? 'break-inside:avoid' : ''].filter(Boolean)
    const cells = row.cells.map((cell, cellIndex) => {
      const owner = grid.storedCell(rowIndex, cellIndex)!, margin = documentTableCellMargin(table, cell)
      const borders = documentTableCellBorders(grid, owner)
      const styles = ['background:' + documentTableCellFill(table, cell),
        'vertical-align:' + (cell.style.verticalAlign === 'center' ? 'middle' : cell.style.verticalAlign ?? 'top')]
      for (const side of DOCUMENT_TABLE_BORDER_SIDES) {
        const border = borders[side], kind = border.style === 'dot' ? 'dotted' : border.style === 'dash' ? 'dashed' : 'solid'
        styles.push('padding-' + side + ':' + margin[side] + 'px', 'border-' + side + ':' + border.width + 'px ' + kind + ' ' + border.color)
      }
      const content = documentBlocksHtml(cell.blocks) + documentTableContinuationEntries(cell).map(entry =>
        '<div data-office-table-continuation="v1" data-office-row-offset="' + entry.offset + '">' + documentBlocksHtml(entry.content.blocks) + '</div>').join('')
      return '<td rowspan="' + cell.rowSpan + '" colspan="' + cell.colSpan + '" style="' + escapeHtml(styles.join(';')) + '">' + content + '</td>'
    }).join('')
    return '<tr data-office-row-height-rule="' + properties.heightRule + '" data-office-row-allow-split="' + properties.allowSplit
      + '" style="' + escapeHtml(rowStyles.join(';')) + '">' + cells + '</tr>'
  })
  const header = table.headerRowCount ?? 0
  const groupable = header > 0 && !grid.anchors.some(owner => owner.row < header && owner.row + owner.rowSpan > header)
  const content = groupable ? '<thead>' + rows.slice(0, header).join('') + '</thead><tbody>' + rows.slice(header).join('') + '</tbody>'
    : '<tbody>' + rows.join('') + '</tbody>'
  return '<table' + attributes + ' style="' + escapeHtml(tableStyles.join(';')) + '">' + columns + content + '</table>'
}

function documentBlocksHtml(blocks: readonly DocumentBlock[]): string {
  return blocks.map((block, index) => {
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
    if (table) return documentTableHtml(table)
    const image = imageForBlock(block)
    if (image) return `<p><img src="${escapeHtml(image.src)}" alt="${escapeHtml(image.alt)}"${image.width ? ` width="${image.width}"` : ''}${image.height ? ` height="${image.height}"` : ''}></p>`
    if (block.type === 'pageBreak') return '<hr data-office-page-break="true" style="break-after:page">'
    if (block.type === 'columnBreak') return '<hr data-office-column-break="true">'
    if (block.type === 'sectionBreak') return `<hr data-office-section-break="${block.sectionBreak?.type ?? 'nextPage'}">`
    if (block.type === 'divider') return '<hr>'
    if (block.type === 'checklist') {
      const listId = block.checklist?.listId ?? block.id
      const previous = blocks[index - 1]
      if (previous?.type === 'checklist' && (previous.checklist?.listId ?? previous.id) === listId) return ''
      const items: string[] = []
      for (let cursor = index; cursor < blocks.length; cursor++) {
        const item = blocks[cursor]
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
