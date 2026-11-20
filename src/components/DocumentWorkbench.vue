<script setup lang="ts">
import {
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  Bold,
  Check,
  ChevronLeft,
  ChevronRight,
  Clipboard,
  FileText,
  FileUp,
  Heading1,
  Highlighter,
  IndentDecrease,
  IndentIncrease,
  Italic,
  ImagePlus,
  Link2,
  List,
  ListOrdered,
  Minus,
  Plus,
  Printer,
  Redo2,
  Save,
  Search,
  Strikethrough,
  Underline,
  Undo2,
  ZoomIn,
  ZoomOut,
} from 'lucide-vue-next'
import { computed, nextTick, onMounted, onUnmounted, ref, watch } from 'vue'
import type { DocumentBlock, DocumentBlockType, DocumentModel, DocumentPageSetup, DocumentParagraphStyle, InlineRun } from '../domain/model'

type EditorTab = 'home' | 'insert' | 'layout' | 'view'
type PageView = { id: string; blocks: DocumentBlock[]; setup: DocumentPageSetup }
type DocumentSnapshot = PageView[]

const props = defineProps<{
  modelValue?: DocumentModel | string
  fileName?: string
}>()

const emit = defineEmits<{
  'update:modelValue': [value: DocumentModel]
  'update:fileName': [value: string]
  'document-change': [value: DocumentModel]
  'import-file': [file: File]
  save: [payload: { fileName: string; html: string; model: DocumentModel }]
}>()

const initialHtml = '<h1>在这里开始撰写</h1><p>选择文字后可以调整格式，使用“插入”添加表格、链接或新页面。</p><p><br></p>'
const pages = ref<PageView[]>(props.modelValue && typeof props.modelValue !== 'string'
  ? pagesFromModel(props.modelValue)
  : [createPageModel(typeof props.modelValue === 'string' ? props.modelValue : initialHtml)])
const activePageId = ref(pages.value[0].id)
const activeTab = ref<EditorTab>('home')
const documentTitle = ref(props.fileName || '未命名文档')
const editorCanvasRef = ref<HTMLElement | null>(null)
const fileInputRef = ref<HTMLInputElement | null>(null)
const zoom = ref(100)
const paperSize = ref<'A4' | 'Letter'>('A4')
const orientation = ref<'portrait' | 'landscape'>('portrait')
const showOutline = ref(true)
const isFocusMode = ref(false)
const searchText = ref('')
const savedStatus = ref<'saved' | 'editing'>('saved')
const fontFamily = ref('Arial')
const fontSize = ref('3')
const textColor = ref('#25282d')
const highlightColor = ref('#fff1a8')
const historyPast = ref<DocumentSnapshot[]>([])
const historyFuture = ref<DocumentSnapshot[]>([])
const activeBlockId = ref(pages.value[0]?.blocks[0]?.id || '')
const renderRevision = ref(0)
let savedRange: Range | null = null
let lastEmittedValue: DocumentModel | undefined
let saveTimer: ReturnType<typeof setTimeout> | undefined
let changeTimer: ReturnType<typeof setTimeout> | undefined
let typingTimer: ReturnType<typeof setTimeout> | undefined
let typingCommandOpen = false

const activePageIndex = computed(() => pages.value.findIndex((page) => page.id === activePageId.value))
const activePage = computed(() => pages.value[Math.max(activePageIndex.value, 0)])
const allText = computed(() => pages.value.flatMap((page) => page.blocks).map((block) => htmlToText(blockHtml(block))).join(' '))
const wordCount = computed(() => allText.value.trim().replace(/\s/g, '').length)
const serializedHtml = computed(() => pages.value.map((page) => page.blocks.map((block) => blockHtml(block)).join('')).join('<hr data-office-page-break="true">'))
const pageStyle = computed(() => ({ zoom: zoom.value / 100 }))

watch(() => props.fileName, (value) => {
  if (value) documentTitle.value = value
})

watch(() => props.modelValue, (value) => {
  if (value == null || value === lastEmittedValue || (typeof value !== 'string' && sameModel(value, lastEmittedValue))) return
  const nextPages = typeof value === 'string' ? parseHtmlDocument(value) : pagesFromModel(value)
  if (!nextPages.length) return
  pages.value = nextPages
  renderRevision.value += 1
  activePageId.value = nextPages[0].id
  activeBlockId.value = nextPages[0].blocks[0]?.id || ''
  if (typeof value !== 'string') {
    documentTitle.value = value.name
    syncPageControls(value.page)
  }
  historyPast.value = []
  historyFuture.value = []
  savedStatus.value = 'saved'
}, { immediate: true })

onMounted(() => {
  document.addEventListener('selectionchange', rememberSelection)
})

onUnmounted(() => {
  document.removeEventListener('selectionchange', rememberSelection)
  if (saveTimer) clearTimeout(saveTimer)
  if (changeTimer) clearTimeout(changeTimer)
  if (typingTimer) clearTimeout(typingTimer)
})

function makeId(prefix: string) {
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
}

function htmlToText(html: string) {
  const node = document.createElement('div')
  node.innerHTML = html
  return node.textContent || ''
}

function createBlock(html: string, type: DocumentBlockType = 'paragraph'): DocumentBlock {
  const parsed = document.createElement('div')
  parsed.innerHTML = sanitizeHtml(html)
  const firstTag = parsed.firstElementChild?.tagName.toLowerCase()
  const inferredType: DocumentBlockType = firstTag === 'h1' || firstTag === 'h2' || firstTag === 'h3' ? 'heading' : firstTag === 'blockquote' ? 'quote' : firstTag === 'ul' ? 'bullet' : firstTag === 'ol' ? 'numbered' : firstTag === 'table' ? 'table' : firstTag === 'hr' ? 'divider' : firstTag === 'img' ? 'image' : type
  const level = firstTag === 'h2' ? 2 : firstTag === 'h3' ? 3 : inferredType === 'heading' ? 1 : 0
  const runs = extractInlineRuns(parsed)
  return {
    id: makeId('block'),
    type: inferredType,
    level,
    runs: runs.length ? runs : [emptyRun()],
    style: { align: 'left', indent: 0, lineHeight: 1.5, spaceBefore: 0, spaceAfter: 8 },
    metadata: { html: parsed.innerHTML },
  }
}

function basePageSetup(): DocumentPageSetup {
  return { width: 794, height: 1123, marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72 }
}

function createPageModel(html: string, setup: DocumentPageSetup = basePageSetup()): PageView {
  return {
    id: makeId('page'),
    blocks: parseBlocks(html),
    setup: { ...setup },
  }
}

function parseBlocks(html: string): DocumentBlock[] {
  const root = document.createElement('div')
  root.innerHTML = html || '<p><br></p>'
  const blocks: DocumentBlock[] = []
  root.childNodes.forEach((node) => {
    if (node.nodeType === Node.TEXT_NODE && node.textContent?.trim()) blocks.push(createBlock(`<p>${escapeHtml(node.textContent)}</p>`))
    if (node.nodeType !== Node.ELEMENT_NODE) return
    const element = node as HTMLElement
    blocks.push(createBlock(element.outerHTML))
  })
  return blocks.length ? blocks : [createBlock('<p><br></p>')]
}

function parseHtmlDocument(html: string): PageView[] {
  return html.split(/<hr[^>]*data-office-page-break=["']true["'][^>]*>/i).map((pageHtml) => createPageModel(pageHtml))
}

function escapeHtml(text: string) {
  return text.replace(/[&<>"']/g, (character) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[character] || character)
}

function extractInlineRuns(root: HTMLElement): InlineRun[] {
  const runs: InlineRun[] = []
  const walk = (node: Node, marks: Omit<InlineRun, 'id' | 'text'>) => {
    if (node.nodeType === Node.TEXT_NODE && node.textContent) {
      runs.push({ id: makeId('run'), text: node.textContent, ...marks })
      return
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return
    const element = node as HTMLElement
    const style = element.style
    const legacyFont = element.tagName === 'FONT' ? element as HTMLElement & { color?: string; face?: string; size?: string } : null
    const legacySize = Number(legacyFont?.size || 0)
    const nextMarks: Omit<InlineRun, 'id' | 'text'> = {
      bold: marks.bold || element.tagName === 'B' || element.tagName === 'STRONG' || style.fontWeight === 'bold',
      italic: marks.italic || element.tagName === 'I' || element.tagName === 'EM' || style.fontStyle === 'italic',
      underline: marks.underline || element.tagName === 'U' || style.textDecoration.includes('underline'),
      strike: marks.strike || element.tagName === 'S' || element.tagName === 'STRIKE' || style.textDecoration.includes('line-through'),
      color: style.color || legacyFont?.color || marks.color,
      fontFamily: style.fontFamily || legacyFont?.face || marks.fontFamily,
      fontSize: Number.parseFloat(style.fontSize) || (legacySize > 0 ? 8 + legacySize * 2 : marks.fontSize),
    }
    node.childNodes.forEach((child) => walk(child, nextMarks))
  }
  walk(root, { bold: false, italic: false, underline: false, strike: false, color: '#1f2937', fontFamily: 'Aptos', fontSize: 11 })
  return runs
}

function emptyRun(): InlineRun {
  return { id: makeId('run'), text: '', bold: false, italic: false, underline: false, strike: false, color: '#1f2937', fontFamily: 'Aptos', fontSize: 11 }
}

function defaultPageSetup(): DocumentPageSetup {
  const isLetter = paperSize.value === 'Letter'
  const landscape = orientation.value === 'landscape'
  return {
    width: landscape ? (isLetter ? 1056 : 1123) : (isLetter ? 816 : 794),
    height: landscape ? (isLetter ? 816 : 794) : (isLetter ? 1056 : 1123),
    marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
  }
}

function syncPageControls(setup: DocumentPageSetup) {
  const shortSide = Math.min(setup.width, setup.height)
  paperSize.value = Math.abs(shortSide - 816) < Math.abs(shortSide - 794) ? 'Letter' : 'A4'
  orientation.value = setup.width > setup.height ? 'landscape' : 'portrait'
}

function clonePages(value: PageView[]) {
  return JSON.parse(JSON.stringify(value)) as PageView[]
}

function cloneBlock(value: DocumentBlock): DocumentBlock {
  return JSON.parse(JSON.stringify(value)) as DocumentBlock
}

function sameModel(left: DocumentModel, right?: DocumentModel): boolean {
  return Boolean(right && left.kind === right.kind && left.name === right.name && left.modified === right.modified && left.wordCount === right.wordCount && JSON.stringify(left.page) === JSON.stringify(right.page) && JSON.stringify(left.blocks) === JSON.stringify(right.blocks))
}

function pagesFromModel(value: DocumentModel): PageView[] {
  const setup = { ...value.page }
  const pageViews: PageView[] = []
  let blocks: DocumentBlock[] = []
  let pageId = 'page-start'
  for (const block of value.blocks) {
    if (block.type === 'pageBreak') {
      pageViews.push({ id: pageId, blocks: blocks.length ? blocks : [createBlock('<p><br></p>')], setup })
      blocks = []
      pageId = block.id
    } else {
      blocks.push(cloneBlock(block))
    }
  }
  pageViews.push({ id: pageId, blocks: blocks.length ? blocks : [createBlock('<p><br></p>')], setup })
  return pageViews
}

function modelFromPages(sourcePages = pages.value): DocumentModel {
  const blocks: DocumentBlock[] = []
  sourcePages.forEach((page, index) => {
    blocks.push(...page.blocks.map(cloneBlock))
    if (index < sourcePages.length - 1) {
      blocks.push({ id: sourcePages[index + 1].id, type: 'pageBreak', level: 0, runs: [], style: { align: 'left', indent: 0, lineHeight: 1.5, spaceBefore: 0, spaceAfter: 0 }, metadata: {} })
    }
  })
  return { kind: 'document', name: documentTitle.value, blocks, page: { ...(sourcePages[0]?.setup ?? defaultPageSetup()) }, modified: true, wordCount: wordCount.value }
}

function currentModel(): DocumentModel {
  return modelFromPages()
}

function blockHtml(block: DocumentBlock): string {
  const html = block.metadata.html
  if (typeof html === 'string') return sanitizeHtml(html)
  const text = block.runs.map((run) => {
    let value = escapeHtml(run.text)
    if (run.bold) value = `<strong>${value}</strong>`
    if (run.italic) value = `<em>${value}</em>`
    if (run.underline) value = `<u>${value}</u>`
    if (run.strike) value = `<s>${value}</s>`
    if (run.color !== '#1f2937' || run.fontFamily !== 'Aptos' || run.fontSize !== 11) value = `<span style="color:${escapeHtml(run.color)};font-family:${escapeHtml(run.fontFamily)};font-size:${run.fontSize}pt">${value}</span>`
    return value
  }).join('') || '<br>'
  const tag = block.type === 'heading' ? `h${Math.min(3, Math.max(1, block.level || 1))}` : block.type === 'quote' ? 'blockquote' : block.type === 'divider' ? 'hr' : 'p'
  if (block.type === 'bullet' || block.type === 'numbered') return `<${block.type === 'bullet' ? 'ul' : 'ol'}><li>${text}</li></${block.type === 'bullet' ? 'ul' : 'ol'}>`
  return `<${tag}>${text}</${tag}>`
}

function paperStyle(page: PageView) {
  return {
    width: `${page.setup.width}px`,
    minHeight: `${page.setup.height}px`,
    padding: `${page.setup.marginTop}px ${page.setup.marginRight}px ${page.setup.marginBottom}px ${page.setup.marginLeft}px`,
  }
}

function blockStyle(block: DocumentBlock) {
  return {
    textAlign: block.style.align,
    paddingLeft: `${block.style.indent * 24}px`,
    lineHeight: String(block.style.lineHeight),
    marginTop: `${block.style.spaceBefore}px`,
    marginBottom: `${block.style.spaceAfter}px`,
  }
}

function sanitizeHtml(html: string): string {
  const root = document.createElement('template')
  root.innerHTML = html
  const allowed = new Set(['A', 'B', 'BLOCKQUOTE', 'BR', 'CODE', 'DIV', 'EM', 'FONT', 'H1', 'H2', 'H3', 'HR', 'I', 'IMG', 'LI', 'OL', 'P', 'S', 'SPAN', 'STRIKE', 'STRONG', 'SUB', 'SUP', 'TABLE', 'TBODY', 'TD', 'TH', 'THEAD', 'TR', 'U', 'UL'])
  root.content.querySelectorAll('*').forEach((element) => {
    if (!allowed.has(element.tagName)) {
      element.replaceWith(...Array.from(element.childNodes))
      return
    }
    if (element.tagName === 'FONT') {
      const legacy = element as HTMLElement & { color?: string; face?: string; size?: string }
      const span = document.createElement('span')
      const styles = [legacy.color ? `color:${legacy.color}` : '', legacy.face ? `font-family:${legacy.face}` : '', legacy.size ? `font-size:${8 + Number(legacy.size) * 2}pt` : ''].filter(Boolean)
      if (styles.length) span.setAttribute('style', styles.join(';'))
      span.innerHTML = element.innerHTML
      element.replaceWith(span)
      return
    }
    Array.from(element.attributes).forEach((attribute) => {
      const safe = attribute.name === 'style' || (element.tagName === 'A' && attribute.name === 'href') || (element.tagName === 'IMG' && ['src', 'alt', 'width', 'height'].includes(attribute.name)) || attribute.name.startsWith('data-')
      const unsafeUrl = /^(href|src)$/i.test(attribute.name) && !/^(https?:|mailto:|#|blob:|data:image\/(png|jpeg|gif|webp);base64,)/i.test(attribute.value.trim())
      if (!safe || unsafeUrl) element.removeAttribute(attribute.name)
    })
    if (element.hasAttribute('style')) {
      const allowedStyles = new Set(['background-color', 'border', 'border-collapse', 'color', 'font-family', 'font-size', 'font-style', 'font-weight', 'height', 'line-height', 'margin-bottom', 'margin-left', 'margin-right', 'margin-top', 'padding', 'padding-bottom', 'padding-left', 'padding-right', 'padding-top', 'text-align', 'text-decoration', 'vertical-align', 'width'])
      const style = (element as HTMLElement).style
      Array.from(style).forEach((property) => { if (!allowedStyles.has(property)) style.removeProperty(property) })
      if (!style.length) element.removeAttribute('style')
    }
  })
  return root.innerHTML
}

function emitContent() {
  const model = currentModel()
  lastEmittedValue = model
  emit('update:modelValue', model)
  emit('document-change', model)
}

function rememberSelection() {
  const selection = window.getSelection()
  if (!selection?.rangeCount || !editorCanvasRef.value) return
  const range = selection.getRangeAt(0)
  if (editorCanvasRef.value.contains(range.commonAncestorContainer)) {
    savedRange = range.cloneRange()
    const node = range.commonAncestorContainer.nodeType === Node.ELEMENT_NODE ? range.commonAncestorContainer as Element : range.commonAncestorContainer.parentElement
    const blockNode = node?.closest<HTMLElement>('[data-block-id]')
    const pageNode = node?.closest<HTMLElement>('[data-page-id]')
    activeBlockId.value = blockNode?.dataset.blockId || activeBlockId.value
    if (pageNode?.dataset.pageId) activePageId.value = pageNode.dataset.pageId
  }
}

function restoreSelection() {
  if (!savedRange || !editorCanvasRef.value) return
  const selection = window.getSelection()
  selection?.removeAllRanges()
  selection?.addRange(savedRange)
}

function snapshot() {
  return clonePages(pages.value)
}

function commitCommand(action: () => void, options: { coalesceTyping?: boolean; deferEmit?: boolean } = {}) {
  const coalesceTyping = options.coalesceTyping === true
  if (!coalesceTyping) typingCommandOpen = false
  const recordUndo = !coalesceTyping || !typingCommandOpen
  if (recordUndo) historyPast.value.push(snapshot())
  if (coalesceTyping) {
    typingCommandOpen = true
    if (typingTimer) clearTimeout(typingTimer)
    typingTimer = setTimeout(() => { typingCommandOpen = false }, 650)
  }
  action()
  historyFuture.value = []
  savedStatus.value = 'editing'
  if (options.deferEmit) {
    if (changeTimer) clearTimeout(changeTimer)
    changeTimer = setTimeout(emitContent, 160)
  } else {
    if (changeTimer) clearTimeout(changeTimer)
    emitContent()
  }
  if (saveTimer) clearTimeout(saveTimer)
  saveTimer = setTimeout(() => { savedStatus.value = 'saved' }, 900)
}

function undo() {
  const previous = historyPast.value.pop()
  if (!previous) return
  typingCommandOpen = false
  historyFuture.value.push(snapshot())
  pages.value = previous
  renderRevision.value += 1
  savedRange = null
  activePageId.value = pages.value[Math.max(activePageIndex.value, 0)].id
  savedStatus.value = 'editing'
  emitContent()
}

function redo() {
  const next = historyFuture.value.pop()
  if (!next) return
  typingCommandOpen = false
  historyPast.value.push(snapshot())
  pages.value = next
  renderRevision.value += 1
  savedRange = null
  activePageId.value = pages.value[Math.max(activePageIndex.value, 0)].id
  savedStatus.value = 'editing'
  emitContent()
}

function blockElement(block: DocumentBlock) {
  return editorCanvasRef.value?.querySelector<HTMLElement>(`[data-block-id="${block.id}"]`)
}

function scrollToPage(page: PageView) {
  document.querySelector<HTMLElement>(`[data-page-id="${page.id}"]`)?.scrollIntoView({ behavior: 'smooth', block: 'start' })
}

function syncBlockFromDom(block: DocumentBlock, element: HTMLElement) {
  block.metadata = { ...block.metadata, html: sanitizeHtml(element.innerHTML || '<br>') }
  block.runs = extractInlineRuns(element)
  const firstTag = element.firstElementChild?.tagName.toLowerCase()
  const formattedElement = element.firstElementChild as HTMLElement | null
  const align = formattedElement?.style.textAlign
  if (align === 'left' || align === 'center' || align === 'right' || align === 'justify') block.style.align = align
  if (firstTag === 'h1' || firstTag === 'h2' || firstTag === 'h3') {
    block.type = 'heading'
    block.level = firstTag === 'h2' ? 2 : firstTag === 'h3' ? 3 : 1
  } else if (firstTag === 'ul' || firstTag === 'ol') {
    block.type = firstTag === 'ol' ? 'numbered' : 'bullet'
  } else if (firstTag === 'blockquote') block.type = 'quote'
  else if (firstTag === 'table') block.type = 'table'
  else if (firstTag === 'hr') block.type = 'divider'
  else if (firstTag === 'img') block.type = 'image'
  else block.type = 'paragraph'
}

function runCommand(command: string, value?: string) {
  commitCommand(() => {
    restoreSelection()
    document.execCommand(command, false, value)
    const selection = window.getSelection()
    const touched = pages.value.flatMap((page) => page.blocks).filter((block) => {
      const element = blockElement(block)
      return Boolean(element && selection?.anchorNode && element.contains(selection.anchorNode))
    })
    touched.forEach((block) => {
      const element = blockElement(block)
      if (element) syncBlockFromDom(block, element)
    })
  })
}

function handleBlockInput(event: Event, block: DocumentBlock) {
  const element = event.currentTarget as HTMLElement
  activeBlockId.value = block.id
  commitCommand(() => syncBlockFromDom(block, element), { coalesceTyping: true, deferEmit: true })
}

function applyBlockStyle(key: keyof DocumentParagraphStyle, value: string | number) {
  const block = activePage.value.blocks.find((candidate) => candidate.id === activeBlockId.value)
  if (!block) return
  commitCommand(() => {
    if (key === 'align' && (value === 'left' || value === 'center' || value === 'right' || value === 'justify')) block.style.align = value
    if (key === 'indent' && typeof value === 'number') block.style.indent = Math.max(0, Math.min(8, value))
    if (key === 'lineHeight' && typeof value === 'number') block.style.lineHeight = value
    if (key === 'spaceBefore' && typeof value === 'number') block.style.spaceBefore = value
    if (key === 'spaceAfter' && typeof value === 'number') block.style.spaceAfter = value
  })
}

function changeIndent(direction: -1 | 1) {
  const block = activePage.value.blocks.find((candidate) => candidate.id === activeBlockId.value)
  if (block) applyBlockStyle('indent', block.style.indent + direction)
}

function changePageSetup() {
  const nextSetup = defaultPageSetup()
  commitCommand(() => { pages.value.forEach((page) => { page.setup = { ...nextSetup } }) })
}

function fragmentHtml(fragment: DocumentFragment) {
  const wrapper = document.createElement('div')
  wrapper.appendChild(fragment)
  return wrapper.innerHTML || '<br>'
}

function splitBlock(block: DocumentBlock, element: HTMLElement, range: Range) {
  const beforeRange = range.cloneRange()
  beforeRange.setStartBefore(element.firstChild || element)
  const afterRange = range.cloneRange()
  afterRange.setEndAfter(element.lastChild || element)
  const beforeHtml = fragmentHtml(beforeRange.cloneContents())
  const afterHtml = fragmentHtml(afterRange.cloneContents())
  const index = activePage.value.blocks.findIndex((item) => item.id === block.id)
  const nextBlock = createBlock(afterHtml, block.type)
  nextBlock.style = { ...block.style }
  block.metadata = { ...block.metadata, html: sanitizeHtml(beforeHtml) }
  block.runs = parseRunsFromHtml(beforeHtml)
  activePage.value.blocks.splice(index + 1, 0, nextBlock)
  renderRevision.value += 1
  void nextTick(() => {
    const nextElement = blockElement(nextBlock)
    if (!nextElement) return
    nextElement.focus()
    const caret = document.createRange()
    caret.selectNodeContents(nextElement)
    caret.collapse(true)
    const selection = window.getSelection()
    selection?.removeAllRanges()
    selection?.addRange(caret)
  })
}

function parseRunsFromHtml(html: string) {
  const root = document.createElement('div')
  root.innerHTML = html
  return extractInlineRuns(root)
}

function handleBlockKeydown(event: KeyboardEvent, block: DocumentBlock) {
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
    event.preventDefault()
    saveDocument()
    return
  }
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'z') {
    event.preventDefault()
    if (event.shiftKey) redo()
    else undo()
    return
  }
  if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') {
    event.preventDefault()
    insertPageBreak()
    return
  }
  if (event.key !== 'Enter' || event.shiftKey) return
  if (block.type === 'table') return
  const selection = window.getSelection()
  const element = event.currentTarget as HTMLElement
  if (!selection?.rangeCount || !element.contains(selection.anchorNode)) return
  event.preventDefault()
  commitCommand(() => splitBlock(block, element, selection.getRangeAt(0)))
}

function updateTitle() {
  emit('update:fileName', documentTitle.value)
  savedStatus.value = 'editing'
  emitContent()
}

function saveDocument() {
  emit('save', { fileName: documentTitle.value, html: serializedHtml.value, model: currentModel() })
  savedStatus.value = 'saved'
}

function openImport() {
  fileInputRef.value?.click()
}

function onFileChange(event: Event) {
  const input = event.target as HTMLInputElement
  const file = input.files?.[0]
  if (file) emit('import-file', file)
  input.value = ''
}

function createPage() {
  commitCommand(() => {
    const page = createPageModel('<p><br></p>')
    pages.value.splice(activePageIndex.value + 1, 0, page)
    activePageId.value = page.id
    activeBlockId.value = page.blocks[0].id
  })
  void nextTick(() => {
    scrollToPage(activePage.value)
    blockElement(activePage.value.blocks[0])?.focus()
  })
}

function movePage(direction: -1 | 1) {
  const next = Math.max(0, Math.min(pages.value.length - 1, activePageIndex.value + direction))
  activePageId.value = pages.value[next].id
  activeBlockId.value = pages.value[next].blocks[0]?.id || ''
  scrollToPage(pages.value[next])
}

function setActivePage(page: PageView, shouldScroll = true) {
  activePageId.value = page.id
  activeBlockId.value = page.blocks[0]?.id || ''
  if (shouldScroll) scrollToPage(page)
}

function insertPageBreak() {
  const pageIndex = Math.max(0, activePageIndex.value)
  const page = pages.value[pageIndex]
  if (!page) return

  const blockIndex = page.blocks.findIndex((block) => block.id === activeBlockId.value)
  const block = page.blocks[blockIndex]
  const element = block ? blockElement(block) : undefined
  const selection = window.getSelection()
  const selectedRange = selection?.rangeCount ? selection.getRangeAt(0) : null
  const range = selectedRange && element?.contains(selectedRange.startContainer) && element.contains(selectedRange.endContainer)
    ? selectedRange.cloneRange()
    : savedRange && element?.contains(savedRange.startContainer) && element.contains(savedRange.endContainer)
      ? savedRange.cloneRange()
      : null

  let nextPage: PageView | undefined
  commitCommand(() => {
    let followingBlocks: DocumentBlock[] = []
    let firstNextBlock: DocumentBlock
    if (block && element && range) {
      const beforeRange = range.cloneRange()
      beforeRange.selectNodeContents(element)
      beforeRange.setEnd(range.startContainer, range.startOffset)
      const afterRange = range.cloneRange()
      afterRange.selectNodeContents(element)
      afterRange.setStart(range.endContainer, range.endOffset)
      const beforeHtml = fragmentHtml(beforeRange.cloneContents())
      const afterHtml = fragmentHtml(afterRange.cloneContents())
      firstNextBlock = createBlock(afterHtml, block.type)
      firstNextBlock.style = { ...block.style }
      block.metadata = { ...block.metadata, html: sanitizeHtml(beforeHtml) }
      block.runs = parseRunsFromHtml(beforeHtml)
      followingBlocks = page.blocks.splice(blockIndex + 1)
    } else {
      firstNextBlock = createBlock('<p><br></p>')
    }

    nextPage = {
      id: makeId('page'),
      blocks: [firstNextBlock, ...followingBlocks],
      setup: { ...page.setup },
    }
    pages.value.splice(pageIndex + 1, 0, nextPage)
    renderRevision.value += 1
    activePageId.value = nextPage.id
    activeBlockId.value = firstNextBlock.id
    savedRange = null
  })
  void nextTick(() => {
    if (!nextPage) return
    scrollToPage(nextPage)
    blockElement(nextPage.blocks[0])?.focus()
  })
}

function insertTable() {
  const rows = Number(window.prompt('行数', '3'))
  const columns = Number(window.prompt('列数', '3'))
  if (!Number.isFinite(rows) || !Number.isFinite(columns) || rows < 1 || columns < 1) return
  const safeRows = Math.min(20, Math.floor(rows))
  const safeColumns = Math.min(12, Math.floor(columns))
  const table = `<table><tbody>${Array.from({ length: safeRows }, () => `<tr>${Array.from({ length: safeColumns }, () => '<td><br></td>').join('')}</tr>`).join('')}</tbody></table>`
  commitCommand(() => {
    const block = createBlock(table, 'table')
    activePage.value.blocks.push(block)
    activeBlockId.value = block.id
  })
}

function insertLink() {
  const url = window.prompt('输入链接地址')
  if (url) runCommand('createLink', url)
}

function insertImage() {
  const source = window.prompt('输入图片地址')
  if (!source) return
  const safeSource = sanitizeHtml(`<img src="${escapeHtml(source)}" alt="文档图片">`)
  const block = createBlock(safeSource, 'image')
  commitCommand(() => {
    activePage.value.blocks.push(block)
    activeBlockId.value = block.id
  })
}

function findNext() {
  if (!searchText.value) return
  const root = editorCanvasRef.value
  if (!root) return
  const query = searchText.value.toLocaleLowerCase()
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
  let node = walker.nextNode()
  while (node) {
    const text = node.textContent?.toLocaleLowerCase() || ''
    const index = text.indexOf(query)
    if (index >= 0) {
      const range = document.createRange()
      range.setStart(node, index)
      range.setEnd(node, index + searchText.value.length)
      const selection = window.getSelection()
      selection?.removeAllRanges()
      selection?.addRange(range)
      return
    }
    node = walker.nextNode()
  }
}

function copySelection() {
  document.execCommand('copy')
}

function setZoom(value: number) {
  zoom.value = Math.max(50, Math.min(150, value))
}

function printDocument() {
  window.print()
}

type BlockContentBinding = { html: string; revision: number }

const vBlockContent = {
  mounted(element: HTMLElement, binding: { value: BlockContentBinding }) {
    element.innerHTML = binding.value.html
  },
  updated(element: HTMLElement, binding: { value: BlockContentBinding; oldValue?: BlockContentBinding }) {
    const needsRefresh = binding.oldValue?.revision !== binding.value.revision || document.activeElement !== element
    if (needsRefresh && element.innerHTML !== binding.value.html) element.innerHTML = binding.value.html
  },
}
</script>

<template>
  <section class="document-workbench" :class="{ 'focus-mode': isFocusMode }">
    <header class="document-header">
      <div class="document-mark"><FileText :size="19" /></div>
      <div class="document-name-area">
        <input v-model="documentTitle" class="document-name" aria-label="文档名称" @change="updateTitle" @keydown.enter="($event.target as HTMLInputElement).blur()">
        <div class="save-indicator" :class="savedStatus">
          <Check :size="12" />
          <span>{{ savedStatus === 'saved' ? '已保存' : '未保存' }}</span>
        </div>
      </div>
      <div class="header-actions">
        <button class="header-icon-button" title="打印" aria-label="打印文档" @click="printDocument"><Printer :size="16" /></button>
        <button class="save-button" @click="saveDocument"><Save :size="15" /><span>保存</span></button>
        <button class="import-button" @click="openImport"><FileUp :size="15" /><span>打开文档</span></button>
        <input ref="fileInputRef" class="hidden-input" type="file" accept=".doc,.docx,.odt,.txt" @change="onFileChange">
      </div>
    </header>

    <nav class="ribbon-tabs" aria-label="文档工具栏">
      <button :class="{ selected: activeTab === 'home' }" @click="activeTab = 'home'">开始</button>
      <button :class="{ selected: activeTab === 'insert' }" @click="activeTab = 'insert'">插入</button>
      <button :class="{ selected: activeTab === 'layout' }" @click="activeTab = 'layout'">布局</button>
      <button :class="{ selected: activeTab === 'view' }" @click="activeTab = 'view'">视图</button>
      <div class="ribbon-tab-spacer"></div>
      <label class="search-box">
        <Search :size="14" />
        <input v-model="searchText" placeholder="查找内容" @keydown.enter.prevent="findNext">
        <button v-if="searchText" title="查找下一个" aria-label="查找下一个" @click="findNext"><ChevronRight :size="14" /></button>
      </label>
    </nav>

    <div v-if="activeTab === 'home'" class="ribbon-panel">
      <div class="tool-group">
        <div class="tool-row">
          <button class="tool-button" title="撤销" aria-label="撤销" :disabled="historyPast.length === 0" @mousedown.prevent @click="undo"><Undo2 :size="16" /></button>
          <button class="tool-button" title="重做" aria-label="重做" :disabled="historyFuture.length === 0" @mousedown.prevent @click="redo"><Redo2 :size="16" /></button>
          <button class="tool-button" title="复制" aria-label="复制" @mousedown.prevent @click="copySelection"><Clipboard :size="15" /></button>
        </div>
        <span class="group-caption">剪贴板</span>
      </div>
      <div class="tool-separator"></div>
      <div class="tool-group formatting-group">
        <div class="tool-row">
          <select v-model="fontFamily" class="font-select" aria-label="字体" @mousedown="rememberSelection" @change="runCommand('fontName', fontFamily)">
            <option value="Arial">Arial</option><option value="Calibri">Calibri</option><option value="Georgia">Georgia</option><option value="'Times New Roman'">Times New Roman</option><option value="'Microsoft YaHei'">微软雅黑</option><option value="SimSun">宋体</option>
          </select>
          <select v-model="fontSize" class="size-select" aria-label="字号" @mousedown="rememberSelection" @change="runCommand('fontSize', fontSize)">
            <option value="1">10</option><option value="2">12</option><option value="3">14</option><option value="4">18</option><option value="5">24</option><option value="6">32</option><option value="7">48</option>
          </select>
          <span class="tool-divider"></span>
          <button class="tool-button" title="加粗" aria-label="加粗" @mousedown.prevent @click="runCommand('bold')"><Bold :size="16" /></button>
          <button class="tool-button" title="斜体" aria-label="斜体" @mousedown.prevent @click="runCommand('italic')"><Italic :size="16" /></button>
          <button class="tool-button" title="下划线" aria-label="下划线" @mousedown.prevent @click="runCommand('underline')"><Underline :size="16" /></button>
          <button class="tool-button" title="删除线" aria-label="删除线" @mousedown.prevent @click="runCommand('strikeThrough')"><Strikethrough :size="16" /></button>
          <label class="color-tool" title="字体颜色">
            <span class="color-icon">A</span>
            <input v-model="textColor" type="color" aria-label="字体颜色" @mousedown="rememberSelection" @input="runCommand('foreColor', textColor)">
          </label>
          <label class="color-tool" title="文本高亮">
            <Highlighter :size="15" />
            <input v-model="highlightColor" type="color" aria-label="文本高亮颜色" @mousedown="rememberSelection" @input="runCommand('hiliteColor', highlightColor)">
          </label>
        </div>
        <div class="tool-row">
          <button class="tool-button" title="左对齐" aria-label="左对齐" @mousedown.prevent @click="applyBlockStyle('align', 'left')"><AlignLeft :size="16" /></button>
          <button class="tool-button" title="居中" aria-label="居中" @mousedown.prevent @click="applyBlockStyle('align', 'center')"><AlignCenter :size="16" /></button>
          <button class="tool-button" title="右对齐" aria-label="右对齐" @mousedown.prevent @click="applyBlockStyle('align', 'right')"><AlignRight :size="16" /></button>
          <button class="tool-button" title="两端对齐" aria-label="两端对齐" @mousedown.prevent @click="applyBlockStyle('align', 'justify')"><AlignJustify :size="16" /></button>
          <span class="tool-divider"></span>
          <button class="tool-button" title="项目符号" aria-label="项目符号" @mousedown.prevent @click="runCommand('insertUnorderedList')"><List :size="16" /></button>
          <button class="tool-button" title="编号列表" aria-label="编号列表" @mousedown.prevent @click="runCommand('insertOrderedList')"><ListOrdered :size="16" /></button>
          <button class="tool-button" title="减少缩进" aria-label="减少缩进" :disabled="(activePage.blocks.find((block) => block.id === activeBlockId)?.style.indent ?? 0) === 0" @mousedown.prevent @click="changeIndent(-1)"><IndentDecrease :size="16" /></button>
          <button class="tool-button" title="增加缩进" aria-label="增加缩进" @mousedown.prevent @click="changeIndent(1)"><IndentIncrease :size="16" /></button>
        </div>
        <span class="group-caption">字体与段落</span>
      </div>
      <div class="tool-separator"></div>
      <div class="tool-group styles-group">
        <button class="style-tile" @mousedown.prevent @click="runCommand('formatBlock', 'H1')"><Heading1 :size="18" /><span>标题 1</span></button>
        <button class="style-tile" @mousedown.prevent @click="runCommand('formatBlock', 'P')"><span class="normal-style">Aa</span><span>正文</span></button>
        <span class="group-caption">样式</span>
      </div>
    </div>

    <div v-else-if="activeTab === 'insert'" class="ribbon-panel insert-panel">
      <button class="ribbon-action" @mousedown.prevent @click="insertTable"><span class="table-glyph">▦</span><span>表格</span></button>
      <button class="ribbon-action" @mousedown.prevent @click="insertImage"><ImagePlus :size="17" /><span>图片</span></button>
      <button class="ribbon-action" @mousedown.prevent @click="insertLink"><Link2 :size="17" /><span>链接</span></button>
      <button class="ribbon-action" @mousedown.prevent @click="insertPageBreak"><Plus :size="17" /><span>分页符</span></button>
      <button class="ribbon-action" @mousedown.prevent @click="runCommand('insertHorizontalRule')"><Minus :size="17" /><span>分隔线</span></button>
    </div>

    <div v-else-if="activeTab === 'layout'" class="ribbon-panel layout-panel">
      <label class="layout-control"><span>纸张</span><select v-model="paperSize" @change="changePageSetup"><option>A4</option><option>Letter</option></select></label>
      <label class="layout-control"><span>方向</span><select v-model="orientation" @change="changePageSetup"><option value="portrait">纵向</option><option value="landscape">横向</option></select></label>
      <button class="ribbon-action" @click="insertPageBreak"><Plus :size="17" /><span>插入分页</span></button>
    </div>

    <div v-else class="ribbon-panel view-panel">
      <button class="ribbon-action" @click="showOutline = !showOutline"><FileText :size="17" /><span>{{ showOutline ? '隐藏导航' : '显示导航' }}</span></button>
      <button class="ribbon-action" @click="isFocusMode = !isFocusMode"><span class="focus-glyph">⛶</span><span>{{ isFocusMode ? '退出专注' : '专注模式' }}</span></button>
      <div class="zoom-control"><button class="tool-button" title="缩小" aria-label="缩小" @click="setZoom(zoom - 10)"><ZoomOut :size="16" /></button><input v-model.number="zoom" type="range" min="50" max="150" aria-label="缩放比例"><span>{{ zoom }}%</span><button class="tool-button" title="放大" aria-label="放大" @click="setZoom(zoom + 10)"><ZoomIn :size="16" /></button></div>
    </div>

    <main class="editor-workspace">
      <aside v-if="showOutline && !isFocusMode" class="outline-pane">
        <div class="outline-heading">页面</div>
        <button v-for="(page, index) in pages" :key="page.id" class="page-outline-item" :class="{ active: page.id === activePageId }" @click="setActivePage(page)">
          <span class="page-thumbnail"><span></span><span></span><span></span></span>
          <span>第 {{ index + 1 }} 页</span>
        </button>
        <button class="outline-add-page" title="新建页面" aria-label="新建页面" @click="createPage"><Plus :size="15" /><span>新建页面</span></button>
      </aside>

      <div class="page-stage">
        <div ref="editorCanvasRef" class="page-stack" :style="pageStyle">
          <article
            v-for="page in pages"
            :key="page.id"
            class="paper-page"
            :class="[paperSize.toLowerCase(), orientation, { 'page-active': page.id === activePageId }]"
            :data-page-id="page.id"
            :style="paperStyle(page)"
          >
            <div class="page-editor" role="document" aria-label="文档分页编辑区">
              <div
                v-for="block in page.blocks"
                :key="block.id"
                :data-block-id="block.id"
                class="document-block"
                :class="`block-${block.type}`"
                :style="blockStyle(block)"
                :contenteditable="block.type !== 'divider' && block.type !== 'image'"
                role="textbox"
                aria-multiline="true"
                spellcheck="false"
                v-block-content="{ html: blockHtml(block), revision: renderRevision }"
                @input="handleBlockInput($event, block)"
                @keydown="handleBlockKeydown($event, block)"
                @focus="setActivePage(page, false)"
                @mouseup="rememberSelection"
                @keyup="rememberSelection"
              ></div>
            </div>
          </article>
        </div>
      </div>
    </main>

    <footer class="document-statusbar">
      <div class="status-left">
        <span>第 {{ Math.max(activePageIndex + 1, 1) }} 页，共 {{ pages.length }} 页</span>
        <span class="status-divider"></span>
        <span>{{ wordCount }} 个字符</span>
      </div>
      <div class="status-right">
        <button class="page-step" title="上一页" aria-label="上一页" :disabled="activePageIndex <= 0" @click="movePage(-1)"><ChevronLeft :size="15" /></button>
        <button class="page-step" title="下一页" aria-label="下一页" :disabled="activePageIndex >= pages.length - 1" @click="movePage(1)"><ChevronRight :size="15" /></button>
        <span class="status-divider"></span>
        <button class="page-step" title="缩小" aria-label="缩小" @click="setZoom(zoom - 10)"><Minus :size="14" /></button>
        <input v-model.number="zoom" class="status-zoom" type="range" min="50" max="150" aria-label="缩放比例">
        <button class="page-step" title="放大" aria-label="放大" @click="setZoom(zoom + 10)"><Plus :size="14" /></button>
        <span class="zoom-value">{{ zoom }}%</span>
      </div>
    </footer>
  </section>
</template>

<style scoped>
.document-workbench {
  --line: #e4e7eb;
  --text: #30343b;
  --muted: #7a828e;
  --blue: #3b82f6;
  min-width: 0;
  height: 100%;
  min-height: 560px;
  display: flex;
  flex-direction: column;
  overflow: hidden;
  color: var(--text);
  background: #f3f5f7;
  font: 13px/1.45 Inter, "Microsoft YaHei", "Segoe UI", sans-serif;
  letter-spacing: 0;
}

.document-header {
  min-height: 54px;
  display: flex;
  align-items: center;
  gap: 11px;
  padding: 0 16px;
  border-bottom: 1px solid var(--line);
  background: #fff;
}

.document-mark { width: 30px; height: 30px; display: grid; place-items: center; color: #3182ce; background: #eaf3ff; border-radius: 6px; }
.document-name-area { min-width: 0; flex: 1; display: flex; align-items: center; gap: 12px; }
.document-name { width: min(320px, 45vw); padding: 4px 5px; border: 1px solid transparent; border-radius: 3px; outline: none; color: #252a31; font-size: 14px; font-weight: 600; }
.document-name:hover, .document-name:focus { border-color: #d5dce5; }
.save-indicator { display: inline-flex; align-items: center; gap: 4px; color: #7d8792; font-size: 11px; white-space: nowrap; }
.save-indicator.editing { color: #b7791f; }
.header-actions { display: flex; align-items: center; gap: 7px; }
.header-icon-button, .save-button, .import-button { height: 31px; display: inline-flex; align-items: center; justify-content: center; gap: 7px; padding: 0 10px; border: 1px solid #dfe3e8; border-radius: 4px; color: #3c4652; background: #fff; cursor: pointer; }
.header-icon-button { width: 32px; padding: 0; }
.header-icon-button:hover, .save-button:hover { background: #f6f8fa; }
.import-button { border-color: #3478df; color: #fff; background: #3478df; }
.import-button:hover { background: #2868c7; }
.hidden-input { display: none; }

.ribbon-tabs { min-height: 39px; display: flex; align-items: stretch; gap: 20px; padding: 0 18px; border-bottom: 1px solid var(--line); background: #fff; }
.ribbon-tabs > button { position: relative; padding: 0 3px; border: 0; color: #606874; background: transparent; font: inherit; cursor: pointer; }
.ribbon-tabs > button.selected { color: #2e6fc5; font-weight: 600; }
.ribbon-tabs > button.selected::after { position: absolute; right: 0; bottom: -1px; left: 0; height: 2px; background: #3478df; content: ''; }
.ribbon-tab-spacer { flex: 1; }
.search-box { align-self: center; height: 26px; min-width: 170px; display: flex; align-items: center; gap: 6px; padding: 0 7px; border: 1px solid #e2e6eb; border-radius: 4px; color: #77818d; background: #fafbfc; }
.search-box input { width: 120px; min-width: 0; border: 0; outline: 0; color: #424a54; background: transparent; font: inherit; }
.search-box button { width: 20px; height: 20px; padding: 0; display: grid; place-items: center; border: 0; color: #5b6672; background: transparent; cursor: pointer; }

.ribbon-panel { min-height: 86px; display: flex; align-items: stretch; gap: 12px; padding: 6px 12px 5px; border-bottom: 1px solid var(--line); background: #fff; }
.tool-group { min-width: 0; position: relative; display: flex; flex-direction: column; justify-content: center; gap: 4px; padding: 1px 6px 14px; }
.tool-row { min-height: 27px; display: flex; align-items: center; gap: 3px; }
.group-caption { position: absolute; right: 0; bottom: 0; left: 0; color: #8b929c; font-size: 10px; text-align: center; }
.tool-separator { align-self: center; width: 1px; height: 62px; background: #e7eaee; }
.tool-divider { height: 19px; width: 1px; margin: 0 2px; background: #e5e8ec; }
.tool-button { width: 28px; height: 27px; display: inline-grid; place-items: center; flex: 0 0 auto; padding: 0; border: 1px solid transparent; border-radius: 3px; color: #424b56; background: transparent; cursor: pointer; }
.tool-button:hover, .page-step:hover:not(:disabled) { border-color: #dae1e9; background: #f1f5f9; }
.font-select, .size-select { height: 26px; border: 1px solid #dfe3e8; border-radius: 3px; color: #38414c; background: #fff; font: inherit; }
.font-select { width: 132px; padding: 0 5px; }
.size-select { width: 52px; padding: 0 4px; }
.color-tool { width: 26px; height: 27px; display: inline-flex; position: relative; flex: 0 0 auto; align-items: center; justify-content: center; gap: 1px; border-radius: 3px; color: #3e4853; cursor: pointer; }
.color-tool:hover { background: #f1f5f9; }
.color-tool input { position: absolute; inset: 0; width: 100%; height: 100%; opacity: 0; cursor: pointer; }
.color-icon { position: relative; font-size: 15px; font-weight: 700; }
.color-icon::after { position: absolute; right: 0; bottom: -2px; left: 0; height: 3px; background: v-bind(textColor); content: ''; }
.styles-group { flex-direction: row; align-items: center; gap: 6px; padding-bottom: 14px; }
.style-tile { width: 74px; height: 52px; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 2px; border: 1px solid #e2e6eb; border-radius: 3px; color: #4f5965; background: #fff; font-size: 10px; cursor: pointer; }
.style-tile:hover { border-color: #8eb7ee; background: #f7fbff; }
.normal-style { font-family: Georgia, serif; font-size: 17px; }
.insert-panel, .layout-panel, .view-panel { align-items: center; gap: 8px; padding-left: 16px; }
.ribbon-action { min-width: 68px; height: 57px; display: inline-flex; flex-direction: column; align-items: center; justify-content: center; gap: 5px; padding: 4px 10px; border: 1px solid transparent; border-radius: 4px; color: #48525e; background: transparent; font: inherit; cursor: pointer; }
.ribbon-action:hover { border-color: #dce5ef; background: #f7f9fb; }
.table-glyph, .focus-glyph { font-size: 19px; line-height: 1; }
.layout-control { display: flex; align-items: center; gap: 8px; padding: 0 10px; color: #66717e; }
.layout-control select { height: 30px; min-width: 94px; padding: 0 7px; border: 1px solid #dfe3e8; border-radius: 3px; color: #38414c; background: #fff; font: inherit; }
.zoom-control { display: flex; align-items: center; gap: 9px; margin-left: 16px; color: #66717e; }
.zoom-control input, .status-zoom { width: 122px; accent-color: #367ddd; }

.editor-workspace { min-height: 0; flex: 1; display: flex; overflow: hidden; }
.outline-pane { width: 190px; flex: 0 0 190px; padding: 15px 11px; border-right: 1px solid var(--line); background: #f8f9fb; overflow: auto; }
.outline-heading { margin: 0 0 12px 5px; color: #7a8490; font-size: 11px; font-weight: 600; }
.page-outline-item { width: 100%; display: flex; align-items: center; gap: 10px; padding: 8px; border: 1px solid transparent; border-radius: 4px; color: #4d5661; background: transparent; font: inherit; text-align: left; cursor: pointer; }
.page-outline-item:hover { background: #eef2f6; }
.page-outline-item.active { border-color: #b8d2f4; color: #2d66b1; background: #edf5ff; }
.page-thumbnail { width: 34px; height: 43px; display: flex; flex-direction: column; gap: 3px; padding: 8px 5px; border: 1px solid #d6dce3; background: #fff; box-shadow: 0 1px 2px #0000000d; }
.page-thumbnail span { height: 2px; background: #cbd2da; }
.page-thumbnail span:first-child { width: 80%; height: 4px; background: #aeb9c5; }
.outline-add-page { width: 100%; display: flex; align-items: center; gap: 8px; margin-top: 8px; padding: 8px; border: 0; border-radius: 4px; color: #6a7480; background: transparent; font: inherit; text-align: left; cursor: pointer; }
.outline-add-page:hover { color: #2e70c6; background: #edf3fa; }
.page-stage { min-width: 0; flex: 1; overflow: auto; padding: 28px 24px 42px; background-color: #f0f2f5; background-image: radial-gradient(#d9dde3 0.65px, transparent 0.65px); background-size: 16px 16px; }
.page-stack { width: max-content; margin: 0 auto; transform-origin: top center; }
.paper-page { box-sizing: border-box; width: 794px; min-height: 1123px; padding: 92px 93px; border: 1px solid #e5e7eb; background: #fff; box-shadow: 0 2px 8px #14213216; }
.paper-page.letter { width: 816px; min-height: 1056px; }
.paper-page.landscape { width: 1123px; min-height: 794px; }
.paper-page.landscape.letter { width: 1056px; min-height: 816px; }
.paper-page + .paper-page { margin-top: 28px; }
.paper-page.page-active { box-shadow: 0 3px 12px #14213224; }
.page-editor { min-height: 937px; outline: 0; color: #25282d; font-family: Arial, "Microsoft YaHei", sans-serif; font-size: 14px; line-height: 1.85; overflow-wrap: anywhere; }
.document-block { min-height: 1.6em; outline: 0; white-space: pre-wrap; }
.document-block:focus { border-radius: 2px; box-shadow: 0 0 0 1px #c8ddf8; }
.document-block[contenteditable="false"] { user-select: none; }
.document-block.block-heading :deep(h1), .document-block.block-heading :deep(h2), .document-block.block-heading :deep(h3) { margin-top: 0; }
.document-block.block-divider { min-height: 0; }
.document-block.block-image :deep(img) { display: block; max-width: 100%; height: auto; }
.document-block.block-quote { padding-left: 15px; border-left: 3px solid #bbcadb; color: #647386; }
.paper-page.landscape .page-editor { min-height: 608px; }
.page-editor :deep(h1) { margin: 0 0 22px; font-size: 28px; line-height: 1.35; font-weight: 600; }
.page-editor :deep(h2) { margin: 24px 0 12px; font-size: 20px; line-height: 1.4; font-weight: 600; }
.page-editor :deep(p) { margin: 0 0 12px; }
.page-editor :deep(table) { width: 100%; margin: 14px 0; border-collapse: collapse; }
.page-editor :deep(td), .page-editor :deep(th) { min-width: 60px; padding: 7px 9px; border: 1px solid #aeb7c2; }
.page-editor :deep(ul), .page-editor :deep(ol) { padding-left: 26px; }
.document-statusbar { min-height: 31px; display: flex; align-items: center; justify-content: space-between; gap: 12px; padding: 0 12px; border-top: 1px solid #dfe3e8; color: #69737f; background: #fff; font-size: 11px; }
.status-left, .status-right { display: flex; align-items: center; gap: 10px; }
.status-divider { width: 1px; height: 14px; background: #dfe3e8; }
.page-step { width: 23px; height: 23px; display: grid; place-items: center; padding: 0; border: 1px solid transparent; border-radius: 3px; color: #606b77; background: transparent; cursor: pointer; }
.page-step:disabled { color: #c1c7cf; cursor: default; }
.status-zoom { width: 100px; }
.zoom-value { width: 39px; text-align: right; }
.focus-mode .ribbon-tabs, .focus-mode .ribbon-panel, .focus-mode .document-header { display: none; }
.focus-mode .page-stage { padding-top: 34px; }

@media (max-width: 900px) {
  .document-header { padding: 0 9px; }
  .document-name { width: 160px; }
  .save-button { width: 32px; padding: 0; }
  .save-button span, .import-button span { display: none; }
  .import-button { width: 32px; padding: 0; }
  .outline-pane { width: 140px; flex-basis: 140px; }
  .page-stage { padding: 20px 10px 32px; }
  .formatting-group { overflow-x: auto; }
}

@media (max-width: 600px) {
  .document-workbench { min-height: 480px; }
  .save-indicator { display: none; }
  .ribbon-tabs { gap: 14px; padding: 0 10px; }
  .search-box { min-width: 29px; width: 29px; padding: 0 6px; }
  .search-box input { display: none; }
  .outline-pane { width: 92px; flex-basis: 92px; padding: 12px 5px; }
  .page-outline-item { flex-direction: column; gap: 4px; padding: 5px 2px; font-size: 10px; text-align: center; }
  .outline-add-page { justify-content: center; padding: 7px 2px; font-size: 10px; }
  .outline-add-page span { display: none; }
  .page-stage { padding: 15px 8px 26px; }
  .document-statusbar { padding: 0 6px; }
  .status-left { gap: 6px; }
  .status-right { gap: 3px; }
  .status-zoom { width: 55px; }
}

@media print {
  .document-header, .ribbon-tabs, .ribbon-panel, .outline-pane, .document-statusbar { display: none !important; }
  .document-workbench, .editor-workspace, .page-stage { height: auto; overflow: visible; padding: 0; background: #fff; }
  .page-stack { zoom: 1 !important; }
  .paper-page { border: 0; box-shadow: none; page-break-after: always; }
}
</style>
