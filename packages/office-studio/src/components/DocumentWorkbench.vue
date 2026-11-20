<script setup lang="ts">
import { ChevronRight, FileText, FileUp, Printer, Save, Search, X } from 'lucide-vue-next'
import { Button, Input, InputNumber, Modal, Tooltip } from 'antdv-next'
import { computed, nextTick, onMounted, onUnmounted, ref, shallowRef, watch } from 'vue'
import type { DocumentModel } from '../domain/model'
import type { FDocument } from '../api/facade'
import { createBlankDocument } from '../domain/model'
import DocumentCanvasViewport from './document/DocumentCanvasViewport.vue'
import DocumentContentControlPanel from './document/DocumentContentControlPanel.vue'
import DocumentOutline from './document/DocumentOutline.vue'
import DocumentFooter from './document/DocumentFooter.vue'
import WorkbenchRibbon from '../platform/ribbon/WorkbenchRibbon.vue'
import { activeOutlineIdForBlock, buildDocumentOutline, documentOutlineEnabled, documentOutlinePlacement, type DocumentOutlineItem } from '../features/document/services/document-outline'
import { createDocumentRibbon, documentCommandFromRibbon } from './document/document-ribbon'
import { clampDocumentZoom, documentCharacterCount, documentSelectionCharacterCount, documentToolbarState } from './document/document-toolbar-state'
import { executeDocumentCommand, type DocumentCommand } from '../features/document/commands/document-commands'
import { emptyLayout, type DocumentLayout } from '../features/document/layout/document-layout'
import type { DocumentViewportLocation } from '../features/document/layout/document-viewport'
import { blocksFromHtml, clampPoint, cloneDocument, collapsedSelection, documentHtml, normalizeDocument, runText, runsForTarget, safeUrl, selectedPlainText, selectionCollapsed, textTargets, type DocumentSelection, type RunPatch } from '../features/document/model/document-model'
import { printDocumentModel } from '../features/document/services/document-print'

const props = withDefaults(defineProps<{ modelValue?: DocumentModel | string; fileName?: string; canUndo?: boolean; canRedo?: boolean; documentApi?: FDocument; showHeader?: boolean; showOutline?: boolean }>(), { showHeader: false, showOutline: undefined })
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
const activeTab = ref('home')
const documentTitle = ref(props.fileName || '无标题文档')
const zoom = ref(100)
const paperSize = ref<'A4' | 'Letter'>('A4')
const orientation = ref<'portrait' | 'landscape'>('portrait')
const outlineOverride = ref<boolean>()
const showOutline = computed({ get: () => documentOutlineEnabled(model.value, outlineOverride.value ?? props.showOutline), set: (value: boolean) => { outlineOverride.value = value } })
const isFocusMode = ref(false)
const workbenchRef = ref<HTMLElement | null>(null)
const workbenchWidth = ref(1440)
const visibleBlockId = ref<string>()
const outlineNavigationTarget = ref<string>()
const outlineMode = computed(() => documentOutlinePlacement({ availableWidth: workbenchWidth.value, pageWidth: model.value.page.width, zoom: zoom.value }))
const outlineItems = computed(() => buildDocumentOutline(model.value))
const activeOutlineId = computed(() => activeOutlineIdForBlock(outlineItems.value, model.value, visibleBlockId.value ?? selection.value.focus.blockId))
const showFind = ref(false)
const searchText = ref('')
const activePageIndex = ref(0)
const textColor = ref('#1f2937')
const highlightColor = ref('#fff1a8')
const pendingMarks = shallowRef<RunPatch>({})
const modalOpen = ref(false)
const modalKind = ref<'table' | 'image' | 'link' | 'color' | 'highlight'>('table')
const modalValue = ref('')
const tableRows = ref<number | null>(3)
const tableCols = ref<number | null>(3)
const modalError = ref('')
const notice = ref('')
const contentControlPanelOpen = ref(false)
const pages = computed(() => layout.value.pages)
const wordCount = computed(() => documentCharacterCount(textTargets(model.value).map((target) => runText(runsForTarget(model.value, target))).join('\n')))
const selectedCharacters = computed(() => documentSelectionCharacterCount(model.value, selection.value))
const toolbarState = computed(() => documentToolbarState(model.value, selection.value, pendingMarks.value))
const currentMarks = computed(() => toolbarState.value.marks)
const modalTitle = computed(() => ({ table: '插入表格', image: '插入图片', link: '插入链接', color: '字体颜色', highlight: '文本高亮颜色' })[modalKind.value])
const ribbonTabs = computed(() => createDocumentRibbon({ state: toolbarState.value, canUndo: !!props.canUndo, canRedo: !!props.canRedo, outline: showOutline.value, focus: isFocusMode.value, paper: paperSize.value, orientation: orientation.value, zoom: zoom.value, contentControlsAvailable: !!props.documentApi }))
let resizeObserver: ResizeObserver | undefined
onMounted(() => {
  if (!workbenchRef.value) return
  resizeObserver = new ResizeObserver(([entry]) => { if (entry) workbenchWidth.value = entry.contentRect.width })
  resizeObserver.observe(workbenchRef.value)
})
onUnmounted(() => resizeObserver?.disconnect())

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
    if (outlineNavigationTarget.value !== selection.value.focus.blockId) outlineNavigationTarget.value = undefined
    publish(next)
    void nextTick(() => canvasRef.value?.focus())
  } catch (error) {
    pendingMarks.value = {}
    notify(error instanceof Error ? error.message : String(error))
    void nextTick(() => canvasRef.value?.focus())
  }
}
function onSelection(value: DocumentSelection) { selection.value = value; visibleBlockId.value = value.focus.blockId; if (outlineNavigationTarget.value !== value.focus.blockId) outlineNavigationTarget.value = undefined; pendingMarks.value = {} }
function toggleFormat(field: 'bold' | 'italic' | 'underline') { execute({ id: 'doc.command.format-text', patch: { [field]: currentMarks.value[field] !== true } }) }
function undo() { pendingMarks.value = {}; emit('undo'); void nextTick(() => canvasRef.value?.focus()) }
function redo() { pendingMarks.value = {}; emit('redo'); void nextTick(() => canvasRef.value?.focus()) }
function changePageSetup() {
  const next = cloneDocument(model.value)
  const width = paperSize.value === 'A4' ? 794 : 816
  const height = paperSize.value === 'A4' ? 1123 : 1056
  next.page.width = orientation.value === 'portrait' ? width : height
  next.page.height = orientation.value === 'portrait' ? height : width
  publish(next)
  void nextTick(() => canvasRef.value?.focus())
}
function updateTitle() { emit('update:fileName', documentTitle.value); publish(cloneDocument(model.value)) }
function saveDocument() { emit('save', { fileName: documentTitle.value, html: documentHtml(model.value), model: model.value }) }
function openImport() { fileInputRef.value?.click() }
function onFileChange(event: Event) { const input = event.target as HTMLInputElement; if (input.files?.[0]) emit('import-file', input.files[0]); input.value = '' }
function movePage(direction: -1 | 1) { setActivePage(Math.max(0, Math.min(pages.value.length - 1, activePageIndex.value + direction))) }
function setActivePage(index: number) { activePageIndex.value = index; canvasRef.value?.scrollToPage(index) }
function jumpToHeading(id: string) { const point = { blockId: id, offset: 0 }; outlineNavigationTarget.value = id; onSelection(collapsedSelection(point)); canvasRef.value?.scrollToBlock(id); canvasRef.value?.focus() }
function onViewportChange(location: DocumentViewportLocation) {
  const target = outlineNavigationTarget.value
  if (target) {
    const heading = layout.value.headings.find((candidate) => candidate.id === target)
    // A short page may not scroll at all; keep the explicitly selected heading
    // while it is visible, including the programmatic scroll emitted by navigation.
    if (heading && heading.y >= location.top && heading.y < location.bottom) { visibleBlockId.value = target; return }
    if (location.reason !== 'scroll') return
    outlineNavigationTarget.value = undefined
  }
  visibleBlockId.value = location.firstVisibleBlockId
}
function navigateOutline(item: DocumentOutlineItem) { if (item.blockId) jumpToHeading(item.blockId); else setActivePage(0) }
function insertTable() { modalKind.value = 'table'; modalError.value = ''; modalOpen.value = true }
function insertImage() { modalKind.value = 'image'; modalValue.value = ''; modalError.value = ''; modalOpen.value = true }
function insertLink() { modalKind.value = 'link'; modalValue.value = currentMarks.value.href || ''; modalError.value = ''; modalOpen.value = true }
function confirmModal() {
  if (modalKind.value === 'table') execute({ id: 'doc.command.insert-table', rows: Math.max(1, Math.min(20, tableRows.value || 3)), cols: Math.max(1, Math.min(20, tableCols.value || 3)) })
  else if (modalKind.value === 'color' || modalKind.value === 'highlight') {
    if (!/^#[0-9a-f]{6}$/i.test(modalValue.value)) { modalError.value = '请输入六位十六进制颜色，例如 #2563eb'; return }
    execute({ id: 'doc.command.format-text', patch: { [modalKind.value]: modalValue.value } })
  }
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
  try { await navigator.clipboard.writeText(text) } catch { notify('无法访问剪贴板，请使用 Ctrl+C 复制') }
  canvasRef.value?.focus()
}
function setZoom(value: number) { zoom.value = clampDocumentZoom(value) }
function printDocument() { void printDocumentModel(model.value).catch((error: unknown) => notify(error instanceof Error ? error.message : String(error))) }
function notify(message: string) { notice.value = message; emit('notify', message) }
function onRibbonExecute(id: string, params?: unknown) {
  if (id === 'core.command.undo') { if (props.canUndo) undo(); return }
  if (id === 'core.command.redo') { if (props.canRedo) redo(); return }
  const command = documentCommandFromRibbon(id, params)
  if (command) { execute(command); return }
  switch (id) {
    case 'doc.command.copy': void copySelection(); break
    case 'doc.command.print': printDocument(); break
    case 'doc.operation.open-table': insertTable(); break
    case 'doc.operation.open-image': insertImage(); break
    case 'doc.operation.open-link': insertLink(); break
    case 'doc.operation.open-content-controls': if (props.documentApi) contentControlPanelOpen.value = true; break
    case 'doc.operation.toggle-outline': showOutline.value = !showOutline.value; break
    case 'doc.operation.toggle-focus': isFocusMode.value = !isFocusMode.value; void nextTick(() => canvasRef.value?.focus()); break
    case 'doc.operation.set-zoom': if (typeof params === 'number') { setZoom(params); void nextTick(() => canvasRef.value?.focus()) }; break
    case 'doc.operation.toggle-find': showFind.value = !showFind.value; if (showFind.value) void nextTick(() => workbenchRef.value?.querySelector<HTMLInputElement>('input[aria-label="查找内容"]')?.focus()); break
    case 'doc.operation.custom-color':
      if (params === 'color' || params === 'highlight') { modalKind.value = params; modalValue.value = params === 'color' ? textColor.value : highlightColor.value; modalError.value = ''; modalOpen.value = true }
      break
    case 'doc.command.set-paper-size': if (params === 'A4' || params === 'Letter') { paperSize.value = params; changePageSetup() }; break
    case 'doc.command.set-orientation': if (params === 'portrait' || params === 'landscape') { orientation.value = params; changePageSetup() }; break
  }
}
defineExpose({ save: saveDocument, focus: () => canvasRef.value?.focus() })
</script>

<template>
  <section ref="workbenchRef" class="document-workbench" :class="{ 'focus-mode': isFocusMode }" aria-label="文档工作台">
    <header v-if="showHeader && !isFocusMode" class="document-header">
      <FileText class="document-mark" :size="20" />
      <Input v-model:value="documentTitle" class="document-name" aria-label="文档名称" variant="borderless" @change="updateTitle" @press-enter="($event.target as HTMLInputElement).blur()" />
      <Tooltip title="打印文档"><Button type="text" aria-label="打印文档" @click="printDocument"><Printer :size="16" /></Button></Tooltip>
      <Button @click="saveDocument"><Save :size="15" />保存</Button>
      <Button type="primary" @click="openImport"><FileUp :size="15" />打开文档</Button>
    </header>
    <input ref="fileInputRef" class="hidden-input" type="file" accept=".doc,.docx,.odt,.txt" @change="onFileChange">
    <WorkbenchRibbon v-if="!isFocusMode" v-model="activeTab" class="document-ribbon" :tabs="ribbonTabs" label="文档功能区" @execute="onRibbonExecute">
      <template #tab-extra>
        <label v-if="showFind" class="search-box">
          <Search :size="14" />
          <Input v-model:value="searchText" variant="borderless" placeholder="查找内容" aria-label="查找内容" @press-enter="findNext" @keydown.esc="showFind = false; canvasRef?.focus()" />
          <Tooltip title="查找下一个"><Button type="text" size="small" aria-label="查找下一个" @mousedown.prevent @click="findNext"><ChevronRight :size="14" /></Button></Tooltip>
          <Tooltip title="关闭查找"><Button type="text" size="small" aria-label="关闭查找" @click="showFind = false; canvasRef?.focus()"><X :size="14" /></Button></Tooltip>
        </label>
      </template>
    </WorkbenchRibbon>
    <main class="editor-workspace">
      <DocumentOutline v-if="showOutline && !isFocusMode" :model="model" :active-id="activeOutlineId" :mode="outlineMode" :max-width="256" @navigate="navigateOutline" />
      <DocumentCanvasViewport ref="canvasRef" :model="model" :selection="selection" :zoom="zoom" @command="execute" @toggle-format="toggleFormat" @update:selection="onSelection" @layout="layout = $event" @active-page="activePageIndex = $event" @viewport-change="onViewportChange" @undo="undo" @redo="redo" @save="saveDocument" @notify="notify" />
    </main>
    <DocumentFooter :active-page="activePageIndex" :page-count="pages.length" :character-count="wordCount" :selected-characters="selectedCharacters" :zoom="zoom" :modified="model.modified" :focus-mode="isFocusMode" @navigate="movePage" @zoom="setZoom" @exit-focus="isFocusMode = false" />
    <div v-if="notice" class="document-notice" role="status"><span>{{ notice }}</span><Tooltip title="关闭提示"><Button type="text" aria-label="关闭提示" @click="notice = ''"><X :size="14" /></Button></Tooltip></div>
    <Modal v-model:open="modalOpen" :title="modalTitle" ok-text="确定" cancel-text="取消" :width="440" destroy-on-hidden @ok="confirmModal">
      <div v-if="modalKind === 'table'" class="insert-form"><label>行数<InputNumber v-model:value="tableRows" :min="1" :max="20" :precision="0" /></label><label>列数<InputNumber v-model:value="tableCols" :min="1" :max="20" :precision="0" /></label></div>
      <div v-else-if="modalKind === 'color' || modalKind === 'highlight'" class="insert-form color-form"><label>{{ modalTitle }}<input v-model="modalValue" type="color" :aria-label="modalTitle"></label><Input v-model:value="modalValue" aria-label="颜色值" placeholder="#1f2937" /></div>
      <div v-else class="insert-form"><label>{{ modalKind === 'image' ? '图片地址' : '链接地址' }}<Input v-model:value="modalValue" placeholder="https://" allow-clear @press-enter="confirmModal" /></label><Button v-if="modalKind === 'image'" type="primary" @click="imageInputRef?.click()"><FileUp :size="15" />选择本地图片</Button><input ref="imageInputRef" class="hidden-input" type="file" accept="image/png,image/jpeg,image/gif,image/webp" @change="onImageFile"></div>
      <p v-if="modalError" role="alert" class="insert-error">{{ modalError }}</p>
    </Modal>
    <DocumentContentControlPanel v-model:open="contentControlPanelOpen" :controls="model.contentControls ?? []" :selection="selection" :model="model" :document-api="documentApi" @error="notify" />
  </section>
</template>

<style scoped>
.document-workbench { position: relative; min-width: 0; height: 100%; min-height: 0; display: flex; flex-direction: column; overflow: hidden; color: #30343b; background: #f3f5f7; font: 13px/1.45 "Microsoft YaHei", "Segoe UI", sans-serif; }
.document-header { flex: 0 0 44px; display: flex; align-items: center; gap: 8px; min-width: 0; padding: 0 12px; border-bottom: 1px solid #e4e7eb; background: #fff; }
.document-mark { flex-shrink: 0; color: #3182ce; }
.document-name { min-width: 60px; flex: 1; }
.document-header :deep(.ant-btn) { display: inline-flex; align-items: center; justify-content: center; gap: 6px; height: 28px; }
.hidden-input { display: none; }
.document-ribbon { flex-shrink: 0; }
.editor-workspace { position: relative; min-width: 0; min-height: 0; flex: 1; display: flex; overflow: hidden; }
.search-box { align-self: center; height: 26px; min-width: 120px; max-width: 240px; display: flex; align-items: center; gap: 4px; padding: 0 5px; border: 1px solid #e2e6eb; border-radius: 4px; color: #77818d; background: #fafbfc; }
.search-box :deep(.ant-input) { width: 125px; min-width: 0; padding: 0; color: #424a54; background: transparent; font: inherit; }
.search-box :deep(.ant-btn) { width: 20px; height: 20px; padding: 0; display: grid; place-items: center; color: #5b6672; }
.document-notice { position: absolute; z-index: 40; bottom: 44px; inset-inline: 12px; display: flex; align-items: center; justify-content: space-between; gap: 10px; padding: 8px 12px; border: 1px solid #c9dcf4; border-radius: 4px; color: #4b6482; background: #eef5ff; font-size: 12px; box-shadow: 0 2px 8px #18375912; }
.document-notice :deep(.ant-btn) { display: grid; place-items: center; padding: 0; color: inherit; }
.insert-form { display: flex; flex-direction: column; gap: 16px; padding: 12px 0; }
.insert-form label { display: flex; align-items: center; gap: 12px; }
.insert-form label > * { flex: 1; }
.color-form input[type="color"] { height: 36px; padding: 2px; border: 1px solid #d6dce4; border-radius: 4px; background: #fff; }
.insert-error { color: #c14444; }
@media (max-width: 650px) { .search-box { width: 120px; } .search-box :deep(.ant-input) { width: 60px; } .document-header { padding: 0 6px; } }
@media print { .document-header, .document-ribbon, .document-notice, :deep(.document-outline), :deep(.document-statusbar) { display: none !important; } .document-workbench, .editor-workspace { height: auto; overflow: visible; background: #fff; } }
</style>
