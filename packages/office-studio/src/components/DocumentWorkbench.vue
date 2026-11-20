<script setup lang="ts">
import { AlignCenter, AlignJustify, AlignLeft, AlignRight, Bold, Check, ChevronLeft, ChevronRight, Clipboard, FileText, FileUp, Heading1, Highlighter, IndentDecrease, IndentIncrease, Italic, ImagePlus, Link2, List, ListOrdered, Maximize, Minimize, Minus, Plus, Printer, Redo2, Save, Search, Strikethrough, Table2, Underline, Undo2, X, ZoomIn, ZoomOut } from 'lucide-vue-next'
import { Button, Input, InputNumber, Modal, Segmented, Select, Tooltip } from 'antdv-next'
import { computed, nextTick, ref, shallowRef, watch } from 'vue'
import type { DocumentModel, DocumentParagraphStyle } from '../domain/model'
import type { FDocument } from '../api/facade'
import { createBlankDocument } from '../domain/model'
import DocumentCanvasViewport from './document/DocumentCanvasViewport.vue'
import DocumentContentControlPanel from './document/DocumentContentControlPanel.vue'
import { blockAtPoint, executeDocumentCommand, type DocumentCommand } from '../features/document/commands/document-commands'
import { emptyLayout, type DocumentLayout } from '../features/document/layout/document-layout'
import { blocksFromHtml, clampPoint, cloneDocument, collapsedSelection, documentHtml, marksAtPoint, normalizeDocument, runText, runsForTarget, safeUrl, selectedPlainText, selectionCollapsed, textTargets, type DocumentSelection, type RunPatch } from '../features/document/model/document-model'
import { printDocumentModel } from '../features/document/services/document-print'

type EditorTab = 'home' | 'insert' | 'layout' | 'view'
const props = defineProps<{ modelValue?: DocumentModel | string; fileName?: string; canUndo?: boolean; canRedo?: boolean; documentApi?: FDocument }>()
const emit = defineEmits<{
  'update:modelValue': [value: DocumentModel]
  'update:fileName': [value: string]
  'document-change': [value: DocumentModel]
  'import-file': [file: File]
  save: [payload: { fileName: string; html: string; model: DocumentModel }]
  undo: []
  redo: []
  notify: [message: string]
}>()
const model = shallowRef<DocumentModel>(normalizeDocument(createBlankDocument()))
const firstTarget = textTargets(model.value)[0]!
const selection = shallowRef<DocumentSelection>(collapsedSelection({ ...firstTarget, offset: 0 }))
const layout = shallowRef<DocumentLayout>(emptyLayout())
const canvasRef = ref<InstanceType<typeof DocumentCanvasViewport> | null>(null)
const fileInputRef = ref<HTMLInputElement | null>(null)
const imageInputRef = ref<HTMLInputElement | null>(null)
const activeTab = ref<EditorTab>('home')
const documentTitle = ref(props.fileName || '无标题文档')
const zoom = ref(100)
const paperSize = ref<'A4' | 'Letter'>('A4')
const orientation = ref<'portrait' | 'landscape'>('portrait')
const showOutline = ref(true)
const isFocusMode = ref(false)
const searchText = ref('')
const activePageIndex = ref(0)
const fontFamily = ref('Arial')
const fontSize = ref('11')
const textColor = ref('#1f2937')
const highlightColor = ref('#fff1a8')
const pendingMarks = shallowRef<RunPatch>({})
const modalOpen = ref(false)
const modalKind = ref<'table' | 'image' | 'link'>('table')
const modalValue = ref('')
const tableRows = ref<number | null>(3)
const tableCols = ref<number | null>(3)
const modalError = ref('')
const notice = ref('')
const contentControlPanelOpen = ref(false)
const savedStatus = computed(() => model.value.modified ? 'editing' : 'saved')
const pages = computed(() => layout.value.pages)
const activeBlockId = computed(() => selection.value.focus.blockId)
const activeBlock = computed(() => blockAtPoint(model.value, selection.value.focus))
const wordCount = computed(() => textTargets(model.value).reduce((count, target) => count + runText(runsForTarget(model.value, target)).replace(/\s/g, '').length, 0))
const currentMarks = computed(() => ({ ...marksAtPoint(model.value, selection.value.focus), ...pendingMarks.value }))
const outline = computed(() => layout.value.headings)
const modalTitle = computed(() => modalKind.value === 'table' ? '插入表格' : modalKind.value === 'image' ? '插入图片' : '插入链接')
const editorTabs = [
  { label: '开始', value: 'home' },
  { label: '插入', value: 'insert' },
  { label: '布局', value: 'layout' },
  { label: '视图', value: 'view' },
]
const fontOptions = [
  { label: 'Arial', value: 'Arial' },
  { label: 'Calibri', value: 'Calibri' },
  { label: 'Georgia', value: 'Georgia' },
  { label: 'Times New Roman', value: "'Times New Roman'" },
  { label: '微软雅黑', value: "'Microsoft YaHei'" },
  { label: '宋体', value: 'SimSun' },
]
const fontSizeOptions = [10, 11, 12, 14, 18, 24, 32, 48].map((value) => ({ label: String(value), value: String(value) }))
const paperSizeOptions = ['A4', 'Letter'].map((value) => ({ label: value, value }))
const orientationOptions = [{ label: '纵向', value: 'portrait' }, { label: '横向', value: 'landscape' }]

watch(() => props.fileName, (value) => { if (value) documentTitle.value = value })
watch(() => props.modelValue, (value) => {
  if (value == null) return
  // A parent history undo/redo can legitimately return to a previously
  // emitted snapshot. Skip only when the local model is already at that
  // snapshot; comparing against lastEmitted alone would leave the canvas on
  // the undo state and make redo appear to do nothing.
  if (typeof value !== 'string' && JSON.stringify(value) === JSON.stringify(model.value)) return
  const next = typeof value === 'string' ? { ...createBlankDocument(), blocks: blocksFromHtml(value) } : value
  model.value = normalizeDocument(next)
  documentTitle.value = next.name
  selection.value = { anchor: clampPoint(model.value, selection.value.anchor), focus: clampPoint(model.value, selection.value.focus) }
  paperSize.value = Math.abs(Math.min(next.page.width, next.page.height) - 816) < 11 ? 'Letter' : 'A4'
  orientation.value = next.page.width > next.page.height ? 'landscape' : 'portrait'
}, { immediate: true })
watch(currentMarks, (marks) => {
  fontFamily.value = marks.fontFamily || 'Arial'
  fontSize.value = String(marks.fontSize || 11)
  textColor.value = marks.color?.startsWith('#') ? marks.color : '#1f2937'
  highlightColor.value = marks.highlight?.startsWith('#') ? marks.highlight : '#fff1a8'
}, { immediate: true })

function publish(next: DocumentModel) {
  next.name = documentTitle.value
  next.modified = true
  next.wordCount = textTargets(next).reduce((count, target) => count + runText(runsForTarget(next, target)).replace(/\s/g, '').length, 0)
  model.value = next
  emit('update:modelValue', next)
  emit('document-change', next)
}
function execute(command: DocumentCommand) {
  if (command.id === 'doc.command.format-text' && selectionCollapsed(selection.value)) {
    pendingMarks.value = { ...pendingMarks.value, ...command.patch }
    canvasRef.value?.focus()
    return
  }
  try {
    const next = cloneDocument(model.value)
    const nextCommand = command.id === 'doc.command.insert-text' ? { ...command, marks: { ...pendingMarks.value, ...command.marks } } : command
    selection.value = executeDocumentCommand(next, selection.value, nextCommand)
    publish(next)
    void nextTick(() => canvasRef.value?.focus())
  } catch (error) {
    pendingMarks.value = {}
    notify(error instanceof Error ? error.message : String(error))
    void nextTick(() => canvasRef.value?.focus())
  }
}
function onSelection(value: DocumentSelection) { selection.value = value; pendingMarks.value = {} }
function undo() { pendingMarks.value = {}; emit('undo') }
function redo() { pendingMarks.value = {}; emit('redo') }
function rememberSelection() { /* Logical selection remains independent of toolbar DOM focus. */ }
function runCommand(command: string, value?: string) {
  const flag = command === 'bold' ? 'bold' : command === 'italic' ? 'italic' : command === 'underline' ? 'underline' : command === 'strikeThrough' ? 'strike' : undefined
  if (flag) execute({ id: 'doc.command.format-text', patch: { [flag]: !currentMarks.value[flag] } })
  else if (command === 'fontName') execute({ id: 'doc.command.format-text', patch: { fontFamily: value || 'Arial' } })
  else if (command === 'fontSize') execute({ id: 'doc.command.format-text', patch: { fontSize: Number(value) || 11 } })
  else if (command === 'foreColor') execute({ id: 'doc.command.format-text', patch: { color: value } })
  else if (command === 'hiliteColor') execute({ id: 'doc.command.format-text', patch: { highlight: value } })
  else if (command === 'formatBlock') execute({ id: 'doc.command.set-block-type', type: value === 'P' ? 'paragraph' : 'heading', level: Number(value?.slice(1)) || 1 })
  else if (command === 'insertUnorderedList' || command === 'insertOrderedList') execute({ id: 'doc.command.set-block-type', type: command === 'insertOrderedList' ? 'numbered' : 'bullet' })
  else if (command === 'insertHorizontalRule') execute({ id: 'doc.command.insert-divider' })
}
function applyBlockStyle(key: keyof DocumentParagraphStyle, value: string | number) {
  if (key === 'align' && ['left', 'center', 'right', 'justify'].includes(String(value))) execute({ id: 'doc.command.format-paragraph', patch: { align: value as DocumentParagraphStyle['align'] } })
  else if (key !== 'align' && typeof value === 'number') execute({ id: 'doc.command.format-paragraph', patch: { [key]: value } })
}
function changeIndent(direction: -1 | 1) { applyBlockStyle('indent', Math.max(0, Math.min(20, (activeBlock.value?.style.indent || 0) + direction))) }
function changePageSetup() {
  const next = cloneDocument(model.value)
  const width = paperSize.value === 'A4' ? 794 : 816
  const height = paperSize.value === 'A4' ? 1123 : 1056
  next.page.width = orientation.value === 'portrait' ? width : height
  next.page.height = orientation.value === 'portrait' ? height : width
  publish(next)
}
function updateTitle() { emit('update:fileName', documentTitle.value); publish(cloneDocument(model.value)) }
function setActiveTab(value: string | number) {
  if (value === 'home' || value === 'insert' || value === 'layout' || value === 'view') activeTab.value = value
}
function updateFontFamily(value: string) { fontFamily.value = value; runCommand('fontName', value) }
function updateFontSize(value: string | number) { fontSize.value = String(value); runCommand('fontSize', String(value)) }
function saveDocument() { emit('save', { fileName: documentTitle.value, html: documentHtml(model.value), model: model.value }) }
function openImport() { fileInputRef.value?.click() }
function onFileChange(event: Event) { const input = event.target as HTMLInputElement; if (input.files?.[0]) emit('import-file', input.files[0]); input.value = '' }
function createPage() { execute({ id: 'doc.command.insert-page-break' }) }
function insertPageBreak() { createPage() }
function movePage(direction: -1 | 1) { setActivePage(Math.max(0, Math.min(pages.value.length - 1, activePageIndex.value + direction))) }
function setActivePage(index: number) { activePageIndex.value = index; canvasRef.value?.scrollToPage(index) }
function jumpToHeading(id: string) { const point = { blockId: id, offset: 0 }; onSelection(collapsedSelection(point)); canvasRef.value?.scrollToBlock(id); canvasRef.value?.focus() }
function insertTable() { modalKind.value = 'table'; modalError.value = ''; modalOpen.value = true }
function insertImage() { modalKind.value = 'image'; modalValue.value = ''; modalError.value = ''; modalOpen.value = true }
function insertLink() { modalKind.value = 'link'; modalValue.value = currentMarks.value.href || ''; modalError.value = ''; modalOpen.value = true }
function confirmModal() {
  if (modalKind.value === 'table') execute({ id: 'doc.command.insert-table', rows: Math.max(1, Math.min(20, tableRows.value || 3)), cols: Math.max(1, Math.min(20, tableCols.value || 3)) })
  else {
    const value = safeUrl(modalValue.value)
    if (!value) { modalError.value = '请输入有效的 http 或 https 地址'; return }
    if (modalKind.value === 'link') execute({ id: 'doc.command.format-text', patch: { href: value, underline: true, color: '#2563b6' } })
    else execute({ id: 'doc.command.insert-image', image: { src: value, width: 320, height: 0, alt: '图片' } })
  }
  modalOpen.value = false
}
async function onImageFile(event: Event) {
  const input = event.target as HTMLInputElement
  const file = input.files?.[0]
  input.value = ''
  if (!file) return
  const source = await new Promise<string>((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(String(reader.result)); reader.onerror = () => reject(reader.error); reader.readAsDataURL(file) }).catch(() => '')
  if (!source) { notify('读取图片失败'); return }
  const image = new Image()
  image.onload = () => { execute({ id: 'doc.command.insert-image', image: { src: source, alt: file.name, width: Math.min(500, image.naturalWidth), height: Math.min(500, image.naturalWidth) * image.naturalHeight / image.naturalWidth } }); modalOpen.value = false }
  image.onerror = () => notify('无法读取该图片')
  image.src = source
}
function findNext() {
  if (!searchText.value) return
  const targets = textTargets(model.value)
  const current = targets.findIndex((target) => target.blockId === selection.value.focus.blockId && target.row === selection.value.focus.row && target.col === selection.value.focus.col)
  for (let step = 0; step <= targets.length; step++) {
    const index = (Math.max(0, current) + step) % targets.length
    const target = targets[index]!
    const text = runText(runsForTarget(model.value, target)).toLocaleLowerCase()
    const start = step === 0 ? selection.value.focus.offset : 0
    const offset = text.indexOf(searchText.value.toLocaleLowerCase(), start)
    if (offset >= 0) { onSelection({ anchor: { ...target, offset }, focus: { ...target, offset: offset + searchText.value.length } }); canvasRef.value?.scrollToBlock(target.blockId); canvasRef.value?.focus(); return }
  }
  notify('未找到匹配内容')
}
async function copySelection() {
  const text = selectedPlainText(model.value, selection.value)
  if (!text) return
  try { await navigator.clipboard.writeText(text) } catch { canvasRef.value?.focus(); notify('无法访问剪贴板，请使用 Ctrl+C 复制') }
}
function setZoom(value: number) { zoom.value = Math.max(10, Math.min(400, value)) }
function printDocument() { void printDocumentModel(model.value).catch((error: unknown) => notify(error instanceof Error ? error.message : String(error))) }
function notify(message: string) { notice.value = message; emit('notify', message) }
</script>

<template>
  <section class="document-workbench" :class="{ 'focus-mode': isFocusMode }">
    <header class="document-header">
      <div class="document-mark"><FileText :size="19" /></div>
      <div class="document-name-area">
        <Input v-model:value="documentTitle" class="document-name" aria-label="文档名称" variant="borderless" @change="updateTitle" @press-enter="($event.target as HTMLInputElement).blur()" />
        <div class="save-indicator" :class="savedStatus">
          <Check :size="12" />
          <span>{{ savedStatus === 'saved' ? '已保存' : '未保存' }}</span>
        </div>
      </div>
      <div class="header-actions">
        <Tooltip title="打印文档"><Button class="header-icon-button" type="text" aria-label="打印文档" @click="printDocument"><Printer :size="16" /></Button></Tooltip>
        <Button class="save-button" type="default" @click="saveDocument"><Save :size="15" /><span>保存</span></Button>
        <Button class="import-button" type="primary" @click="openImport"><FileUp :size="15" /><span>打开文档</span></Button>
        <input ref="fileInputRef" class="hidden-input" type="file" accept=".doc,.docx,.odt,.txt" @change="onFileChange">
      </div>
    </header>

    <nav class="ribbon-tabs" aria-label="文档工具栏">
      <Segmented class="ribbon-segmented" :value="activeTab" :options="editorTabs" @change="setActiveTab" />
      <div class="ribbon-tab-spacer"></div>
      <label class="search-box">
        <Search :size="14" />
        <Input v-model:value="searchText" variant="borderless" placeholder="查找内容" aria-label="查找内容" @press-enter="findNext" />
        <Tooltip v-if="searchText" title="查找下一个"><Button type="text" size="small" aria-label="查找下一个" @click="findNext"><ChevronRight :size="14" /></Button></Tooltip>
      </label>
    </nav>

    <div v-if="activeTab === 'home'" class="ribbon-panel">
      <div class="tool-group">
        <div class="tool-row">
          <Tooltip title="撤销"><Button class="tool-button" type="text" aria-label="撤销" :disabled="!canUndo" @mousedown.prevent @click="undo"><Undo2 :size="16" /></Button></Tooltip>
          <Tooltip title="重做"><Button class="tool-button" type="text" aria-label="重做" :disabled="!canRedo" @mousedown.prevent @click="redo"><Redo2 :size="16" /></Button></Tooltip>
          <Tooltip title="复制"><Button class="tool-button" type="text" aria-label="复制" @mousedown.prevent @click="copySelection"><Clipboard :size="15" /></Button></Tooltip>
        </div>
        <span class="group-caption">剪贴板</span>
      </div>
      <div class="tool-separator"></div>
      <div class="tool-group formatting-group">
        <div class="tool-row">
          <Select v-model:value="fontFamily" class="font-select" aria-label="字体" :options="fontOptions" :show-search="true" :filter-option="true" @mousedown="rememberSelection" @change="updateFontFamily" />
          <Select v-model:value="fontSize" class="size-select" aria-label="字号" :options="fontSizeOptions" @mousedown="rememberSelection" @change="updateFontSize" />
          <span class="tool-divider"></span>
          <Tooltip title="加粗"><Button class="tool-button" type="text" aria-label="加粗" :class="{ active: currentMarks.bold }" @mousedown.prevent @click="runCommand('bold')"><Bold :size="16" /></Button></Tooltip>
          <Tooltip title="斜体"><Button class="tool-button" type="text" aria-label="斜体" :class="{ active: currentMarks.italic }" @mousedown.prevent @click="runCommand('italic')"><Italic :size="16" /></Button></Tooltip>
          <Tooltip title="下划线"><Button class="tool-button" type="text" aria-label="下划线" :class="{ active: currentMarks.underline }" @mousedown.prevent @click="runCommand('underline')"><Underline :size="16" /></Button></Tooltip>
          <Tooltip title="删除线"><Button class="tool-button" type="text" aria-label="删除线" :class="{ active: currentMarks.strike }" @mousedown.prevent @click="runCommand('strikeThrough')"><Strikethrough :size="16" /></Button></Tooltip>
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
          <Tooltip title="左对齐"><Button class="tool-button" type="text" aria-label="左对齐" @mousedown.prevent @click="applyBlockStyle('align', 'left')"><AlignLeft :size="16" /></Button></Tooltip>
          <Tooltip title="居中"><Button class="tool-button" type="text" aria-label="居中" @mousedown.prevent @click="applyBlockStyle('align', 'center')"><AlignCenter :size="16" /></Button></Tooltip>
          <Tooltip title="右对齐"><Button class="tool-button" type="text" aria-label="右对齐" @mousedown.prevent @click="applyBlockStyle('align', 'right')"><AlignRight :size="16" /></Button></Tooltip>
          <Tooltip title="两端对齐"><Button class="tool-button" type="text" aria-label="两端对齐" @mousedown.prevent @click="applyBlockStyle('align', 'justify')"><AlignJustify :size="16" /></Button></Tooltip>
          <span class="tool-divider"></span>
          <Tooltip title="项目符号"><Button class="tool-button" type="text" aria-label="项目符号" @mousedown.prevent @click="runCommand('insertUnorderedList')"><List :size="16" /></Button></Tooltip>
          <Tooltip title="编号列表"><Button class="tool-button" type="text" aria-label="编号列表" @mousedown.prevent @click="runCommand('insertOrderedList')"><ListOrdered :size="16" /></Button></Tooltip>
          <Tooltip title="减少缩进"><Button class="tool-button" type="text" aria-label="减少缩进" :disabled="(activeBlock?.style.indent ?? 0) === 0" @mousedown.prevent @click="changeIndent(-1)"><IndentDecrease :size="16" /></Button></Tooltip>
          <Tooltip title="增加缩进"><Button class="tool-button" type="text" aria-label="增加缩进" @mousedown.prevent @click="changeIndent(1)"><IndentIncrease :size="16" /></Button></Tooltip>
        </div>
        <span class="group-caption">字体与段落</span>
      </div>
      <div class="tool-separator"></div>
      <div class="tool-group styles-group">
        <Button class="style-tile" type="text" @mousedown.prevent @click="runCommand('formatBlock', 'H1')"><Heading1 :size="18" /><span>标题 1</span></Button>
        <Button class="style-tile" type="text" @mousedown.prevent @click="runCommand('formatBlock', 'P')"><span class="normal-style">Aa</span><span>正文</span></Button>
        <span class="group-caption">样式</span>
      </div>
    </div>

    <div v-else-if="activeTab === 'insert'" class="ribbon-panel insert-panel">
      <Button class="ribbon-action" type="text" @mousedown.prevent @click="insertTable"><Table2 :size="17" /><span>表格</span></Button>
      <Button class="ribbon-action" type="text" @mousedown.prevent @click="insertImage"><ImagePlus :size="17" /><span>图片</span></Button>
      <Button class="ribbon-action" type="text" @mousedown.prevent @click="insertLink"><Link2 :size="17" /><span>链接</span></Button>
      <Button class="ribbon-action" type="text" @mousedown.prevent @click="insertPageBreak"><Plus :size="17" /><span>分页符</span></Button>
      <Button class="ribbon-action" type="text" @mousedown.prevent @click="runCommand('insertHorizontalRule')"><Minus :size="17" /><span>分隔线</span></Button>
      <Button class="ribbon-action" type="text" @mousedown.prevent @click="contentControlPanelOpen = true"><FileText :size="17" /><span>内容控件</span></Button>
    </div>

    <div v-else-if="activeTab === 'layout'" class="ribbon-panel layout-panel">
      <label class="layout-control"><span>纸张</span><Select v-model:value="paperSize" :options="paperSizeOptions" aria-label="纸张" @change="changePageSetup" /></label>
      <label class="layout-control"><span>方向</span><Select v-model:value="orientation" :options="orientationOptions" aria-label="方向" @change="changePageSetup" /></label>
      <Button class="ribbon-action" type="text" @click="insertPageBreak"><Plus :size="17" /><span>插入分页</span></Button>
    </div>

    <div v-else class="ribbon-panel view-panel">
      <Button class="ribbon-action" type="text" @click="showOutline = !showOutline"><FileText :size="17" /><span>{{ showOutline ? '隐藏导航' : '显示导航' }}</span></Button>
      <Button class="ribbon-action" type="text" @click="isFocusMode = !isFocusMode"><component :is="isFocusMode ? Minimize : Maximize" :size="17" /><span>{{ isFocusMode ? '退出专注' : '专注模式' }}</span></Button>
      <div class="zoom-control"><Tooltip title="缩小"><Button class="tool-button" type="text" aria-label="缩小" @click="setZoom(zoom - 10)"><ZoomOut :size="16" /></Button></Tooltip><input v-model.number="zoom" type="range" min="10" max="400" aria-label="缩放比例"><span>{{ zoom }}%</span><Tooltip title="放大"><Button class="tool-button" type="text" aria-label="放大" @click="setZoom(zoom + 10)"><ZoomIn :size="16" /></Button></Tooltip></div>
    </div>

    <main class="editor-workspace">
      <aside v-if="showOutline && !isFocusMode" class="outline-pane">
        <div class="outline-heading">文档大纲</div>
        <Button v-for="heading in outline" :key="heading.id" class="page-outline-item" type="text" :class="{ active: heading.id === activeBlockId }" :style="{ paddingLeft: (8 + (heading.level - 1) * 12) + 'px' }" @click="jumpToHeading(heading.id)">{{ heading.text }}</Button>
        <p v-if="!outline.length" class="outline-empty">文档中暂无标题</p>
        <div class="outline-heading page-nav-heading">页面</div>
        <Button v-for="(page, index) in pages" :key="page.index" class="page-outline-item" type="text" :class="{ active: index === activePageIndex }" @click="setActivePage(index)"><span class="page-thumbnail"><span></span><span></span><span></span></span><span>第 {{ index + 1 }} 页</span></Button>
        <Button class="outline-add-page" type="text" title="新建页面" aria-label="新建页面" @click="createPage"><Plus :size="15" /><span>新建页面</span></Button>
      </aside>
      <DocumentCanvasViewport ref="canvasRef" :model="model" :selection="selection" :zoom="zoom" @command="execute" @update:selection="onSelection" @layout="layout = $event" @active-page="activePageIndex = $event" @undo="undo" @redo="redo" @save="saveDocument" @notify="notify" />
    </main>

    <div v-if="notice" class="document-notice" role="status"><span>{{ notice }}</span><Tooltip title="关闭提示"><Button type="text" aria-label="关闭提示" @click="notice = ''"><X :size="14" /></Button></Tooltip></div>
    <Modal v-model:open="modalOpen" :title="modalTitle" ok-text="确定" cancel-text="取消" :width="440" destroy-on-hidden @ok="confirmModal">
      <div v-if="modalKind === 'table'" class="insert-form"><label>行数<InputNumber v-model:value="tableRows" :min="1" :max="20" :precision="0" /></label><label>列数<InputNumber v-model:value="tableCols" :min="1" :max="20" :precision="0" /></label></div>
      <div v-else class="insert-form"><label>{{ modalKind === 'image' ? '图片地址' : '链接地址' }}<Input v-model:value="modalValue" placeholder="https://" allow-clear @press-enter="confirmModal" /></label><Button v-if="modalKind === 'image'" class="import-button" type="primary" @click="imageInputRef?.click()"><FileUp :size="15" />选择本地图片</Button><input ref="imageInputRef" class="hidden-input" type="file" accept="image/png,image/jpeg,image/gif,image/webp" @change="onImageFile"></div>
      <p v-if="modalError" role="alert" class="insert-error">{{ modalError }}</p>
    </Modal>
    <DocumentContentControlPanel
      v-model:open="contentControlPanelOpen"
      :controls="model.contentControls ?? []"
      :selection="selection"
      :model="model"
      :document-api="documentApi"
      @error="notify"
    />
    <footer class="document-statusbar">
      <div class="status-left">
        <span>第 {{ Math.max(activePageIndex + 1, 1) }} 页，共 {{ pages.length }} 页</span>
        <span class="status-divider"></span>
        <span>{{ wordCount }} 个字符</span>
      </div>
      <div class="status-right">
        <Tooltip title="上一页"><Button class="page-step" type="text" aria-label="上一页" :disabled="activePageIndex <= 0" @click="movePage(-1)"><ChevronLeft :size="15" /></Button></Tooltip>
        <Tooltip title="下一页"><Button class="page-step" type="text" aria-label="下一页" :disabled="activePageIndex >= pages.length - 1" @click="movePage(1)"><ChevronRight :size="15" /></Button></Tooltip>
        <span class="status-divider"></span>
        <Tooltip title="缩小"><Button class="page-step" type="text" aria-label="缩小" @click="setZoom(zoom - 10)"><Minus :size="14" /></Button></Tooltip>
        <input v-model.number="zoom" class="status-zoom" type="range" min="10" max="400" aria-label="缩放比例">
        <Tooltip title="放大"><Button class="page-step" type="text" aria-label="放大" @click="setZoom(zoom + 10)"><Plus :size="14" /></Button></Tooltip>
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
.document-name { width: min(320px, 45vw); }
.document-name :deep(.ant-input) { padding: 4px 5px; border-radius: 3px; color: #252a31; font-size: 14px; font-weight: 600; }
.document-name:hover :deep(.ant-input), .document-name:focus-within :deep(.ant-input) { background: #f8fafc; }
.save-indicator { display: inline-flex; align-items: center; gap: 4px; color: #7d8792; font-size: 11px; white-space: nowrap; }
.save-indicator.editing { color: #b7791f; }
.header-actions { display: flex; align-items: center; gap: 7px; }
.header-icon-button.ant-btn, .save-button.ant-btn, .import-button.ant-btn { height: 31px; display: inline-flex; align-items: center; justify-content: center; gap: 7px; padding: 0 10px; border-radius: 4px; }
.header-icon-button.ant-btn { width: 32px; padding: 0; color: #3c4652; }
.save-button.ant-btn { border-color: #dfe3e8; color: #3c4652; }
.import-button.ant-btn { border-color: #3478df; }
.save-button.ant-btn:hover { background: #f6f8fa; }
.hidden-input { display: none; }

.ribbon-tabs { min-height: 39px; display: flex; align-items: center; gap: 12px; padding: 0 18px; border-bottom: 1px solid var(--line); background: #fff; }
.ribbon-segmented { flex: 0 0 auto; }
.ribbon-segmented :deep(.ant-segmented) { padding: 2px; border-radius: 4px; background: #f3f5f7; }
.ribbon-segmented :deep(.ant-segmented-item-label) { min-height: 26px; line-height: 26px; padding: 0 12px; }
.ribbon-tab-spacer { flex: 1; }
.search-box { align-self: center; height: 26px; min-width: 170px; display: flex; align-items: center; gap: 6px; padding: 0 7px; border: 1px solid #e2e6eb; border-radius: 4px; color: #77818d; background: #fafbfc; }
.search-box :deep(.ant-input) { width: 120px; min-width: 0; padding: 0; color: #424a54; background: transparent; font: inherit; }
.search-box :deep(.ant-btn) { width: 20px; height: 20px; padding: 0; display: grid; place-items: center; color: #5b6672; }

.ribbon-panel { min-height: 86px; display: flex; align-items: stretch; gap: 12px; padding: 6px 12px 5px; border-bottom: 1px solid var(--line); background: #fff; }
.tool-group { min-width: 0; position: relative; display: flex; flex-direction: column; justify-content: center; gap: 4px; padding: 1px 6px 14px; }
.tool-row { min-height: 27px; display: flex; align-items: center; gap: 3px; }
.group-caption { position: absolute; right: 0; bottom: 0; left: 0; color: #8b929c; font-size: 10px; text-align: center; }
.tool-separator { align-self: center; width: 1px; height: 62px; background: #e7eaee; }
.tool-divider { height: 19px; width: 1px; margin: 0 2px; background: #e5e8ec; }
.tool-button.ant-btn { width: 28px; min-width: 28px; height: 27px; display: inline-grid; place-items: center; flex: 0 0 auto; padding: 0; border: 1px solid transparent; border-radius: 3px; color: #424b56; }
.tool-button:hover, .page-step:hover:not(:disabled) { border-color: #dae1e9; background: #f1f5f9; }
.font-select { width: 132px; }
.size-select { width: 58px; }
.font-select :deep(.ant-select-selector), .size-select :deep(.ant-select-selector) { height: 26px !important; border-radius: 3px !important; }
.font-select :deep(.ant-select-selection-item), .size-select :deep(.ant-select-selection-item) { line-height: 24px !important; color: #38414c; font: inherit; }
.color-tool { width: 26px; height: 27px; display: inline-flex; position: relative; flex: 0 0 auto; align-items: center; justify-content: center; gap: 1px; border-radius: 3px; color: #3e4853; cursor: pointer; }
.color-tool:hover { background: #f1f5f9; }
.color-tool input { position: absolute; inset: 0; width: 100%; height: 100%; opacity: 0; cursor: pointer; }
.color-icon { position: relative; font-size: 15px; font-weight: 700; }
.color-icon::after { position: absolute; right: 0; bottom: -2px; left: 0; height: 3px; background: v-bind(textColor); content: ''; }
.styles-group { flex-direction: row; align-items: center; gap: 6px; padding-bottom: 14px; }
.style-tile.ant-btn { width: 74px; height: 52px; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 2px; padding: 2px; border: 1px solid #e2e6eb; border-radius: 3px; color: #4f5965; background: #fff; font-size: 10px; }
.style-tile.ant-btn:hover { border-color: #8eb7ee; background: #f7fbff; }
.normal-style { font-family: Georgia, serif; font-size: 17px; }
.insert-panel, .layout-panel, .view-panel { align-items: center; gap: 8px; padding-left: 16px; }
.ribbon-action.ant-btn { min-width: 68px; height: 57px; display: inline-flex; flex-direction: column; align-items: center; justify-content: center; gap: 5px; padding: 4px 10px; border: 1px solid transparent; border-radius: 4px; color: #48525e; background: transparent; font: inherit; }
.ribbon-action.ant-btn:hover { border-color: #dce5ef; background: #f7f9fb; }
.layout-control { display: flex; align-items: center; gap: 8px; padding: 0 10px; color: #66717e; }
.layout-control :deep(.ant-select) { min-width: 94px; color: #38414c; font: inherit; }
.layout-control :deep(.ant-select-selector) { height: 30px !important; border-radius: 3px !important; }
.layout-control :deep(.ant-select-selection-item) { line-height: 28px !important; }
.zoom-control { display: flex; align-items: center; gap: 9px; margin-left: 16px; color: #66717e; }
.zoom-control input, .status-zoom { width: 122px; accent-color: #367ddd; }

.editor-workspace { min-height: 0; flex: 1; display: flex; overflow: hidden; }
.outline-pane { width: 190px; flex: 0 0 190px; padding: 15px 11px; border-right: 1px solid var(--line); background: #f8f9fb; overflow: auto; }
.outline-heading { margin: 0 0 12px 5px; color: #7a8490; font-size: 11px; font-weight: 600; }
.page-outline-item.ant-btn { width: 100%; height: auto; min-height: 34px; display: flex; align-items: center; justify-content: flex-start; gap: 10px; padding: 8px; border: 1px solid transparent; border-radius: 4px; color: #4d5661; background: transparent; font: inherit; text-align: left; white-space: normal; }
.page-outline-item.ant-btn:hover { background: #eef2f6; }
.page-outline-item.ant-btn.active { border-color: #b8d2f4; color: #2d66b1; background: #edf5ff; }
.page-thumbnail { width: 34px; height: 43px; display: flex; flex-direction: column; gap: 3px; padding: 8px 5px; border: 1px solid #d6dce3; background: #fff; box-shadow: 0 1px 2px #0000000d; }
.page-thumbnail span { height: 2px; background: #cbd2da; }
.page-thumbnail span:first-child { width: 80%; height: 4px; background: #aeb9c5; }
.outline-add-page.ant-btn { width: 100%; height: auto; display: flex; align-items: center; justify-content: flex-start; gap: 8px; margin-top: 8px; padding: 8px; border: 0; border-radius: 4px; color: #6a7480; background: transparent; font: inherit; text-align: left; }
.outline-add-page.ant-btn:hover { color: #2e70c6; background: #edf3fa; }
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
.page-step.ant-btn { width: 23px; min-width: 23px; height: 23px; display: grid; place-items: center; padding: 0; border: 1px solid transparent; border-radius: 3px; color: #606b77; }
.page-step.ant-btn:disabled { color: #c1c7cf; cursor: default; }
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
.outline-empty { padding: 0 6px; color: #8b95a3; font-size: 11px; line-height: 1.7; }
.page-nav-heading { margin-top: 24px; }
.document-notice { display: flex; justify-content: space-between; gap: 10px; padding: 7px 14px; color: #4b6482; background: #eef5ff; font-size: 12px; }
.document-notice :deep(.ant-btn) { display: grid; place-items: center; padding: 0; color: inherit; }
.insert-form { display: flex; flex-direction: column; gap: 16px; padding: 12px 0; }
.insert-form label { display: flex; align-items: center; gap: 12px; }
.insert-form label > * { flex: 1; }
.insert-error { color: #c14444; }
.tool-button.active { background: #dfebfb; color: #276ac5; }
</style>
