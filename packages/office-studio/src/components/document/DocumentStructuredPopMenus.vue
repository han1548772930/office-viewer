<script setup lang="ts">
import { computed, shallowRef, type CSSProperties } from 'vue'
import { useI18n } from 'vue-i18n'
import { Sigma } from 'lucide-vue-next'
import { Button, Divider, Drawer, Dropdown, Menu, Popover, Tooltip, type MenuItemType } from 'antdv-next'
import type { DocumentCodeConfig } from '../../domain/model'
import type { RibbonItem } from '../../platform/ribbon/ribbon-schema'
import type { DocumentStructuredCanvasTarget, DocumentStructuredCanvasCommandRequest, DocumentStructuredCanvasOperationRequest } from '../../features/document/controllers/document-structured-canvas'
import { documentStructureContextItems, documentStructureCommandFromMenu, documentStructuredMountedMenuItems,
  type DocumentStructureRibbonContext, type DocumentStructuredParagraphMenuContext } from './document-structure-ribbon'
import { documentStructuredCodeLanguageOptions } from '../../features/document/code-blocks/document-code-languages'
import DocumentCodeControls from './DocumentCodeControls.vue'
import { documentLatexMountedMenuItem } from './document-latex-menu'

interface Control { key?: string; style: CSSProperties; context: DocumentStructuredParagraphMenuContext; empty?: boolean }
const props = defineProps<{ active?: DocumentStructureRibbonContext & { title: string }; toolbarStyle: CSSProperties;
  paragraph?: Control; contentInsert?: Control; tables: Control[]; mobile: boolean; disabled: boolean;
  latexState?: { available: boolean; reason?: string }; floatLatex?: { item: RibbonItem; style: CSSProperties };
  mobileLatex?: { items: RibbonItem[]; style: CSSProperties };
  captureTarget: () => DocumentStructuredCanvasTarget; targetCurrent: (target: DocumentStructuredCanvasTarget) => boolean }>()
const emit = defineEmits<{ command: [request: DocumentStructuredCanvasCommandRequest]; operation: [request: DocumentStructuredCanvasOperationRequest]; focus: [target: DocumentStructuredCanvasTarget] }>()
const { t } = useI18n({ useScope: 'global' })
type MenuKind = 'paragraph' | 'content' | 'table' | 'structure'
interface MenuSession { kind: MenuKind; key?: string; title: string; target: DocumentStructuredCanvasTarget; items: RibbonItem[] }
type CodeControlKind = 'font' | 'language' | 'settings'
interface CodeSession { kind: CodeControlKind; blockId: string; target: DocumentStructuredCanvasTarget; config: DocumentCodeConfig; fontFamily: string; actions: RibbonItem[] }
const menu = shallowRef<MenuSession>()
const code = shallowRef<CodeSession>()
const direct = shallowRef<{ target: DocumentStructuredCanvasTarget; blockId: string }>()
const directLatex = shallowRef<{ target: DocumentStructuredCanvasTarget; itemId: string }>()
const codeTitle = computed(() => code.value?.kind === 'font' ? '等宽字体' : code.value?.kind === 'language' ? '代码语言' : '代码块设置')
const languageTitle = computed(() => documentStructuredCodeLanguageOptions(props.active?.code?.language ?? 'plaintext').find(value => value.value === props.active?.code?.language)?.label ?? '代码语言')
const floatingStyle = computed(() => props.mobile ? { maxWidth: '100%' } : props.toolbarStyle)
const paragraphStyle = computed(() => props.mobile ? undefined : props.paragraph?.style)

function paragraphItems(control: Control): RibbonItem[] {
  const context = control.context, mount = control.empty ? 'quick-insert' : 'quick-edit'
  const choices = documentStructuredMountedMenuItems(context, mount)
  if (control.empty && props.latexState) choices.splice(3, 0, documentLatexMountedMenuItem('insert-menu', props.latexState, t))
  const below = documentStructuredMountedMenuItems(context, 'quick-insert-below')
  const codeActions = props.active?.code ? documentStructureContextItems(props.active).filter(item =>
    ['doc-context-code-language', 'doc-context-code-font', 'doc-context-code-more'].includes(item.id)) : []
  return [...codeActions,
    { id: 'doc-context-structured-insert-current', type: 'menu', title: control.empty ? '插入' : '转换段落', children: choices },
    { id: 'doc-context-structured-insert-below', type: 'menu', title: '在段落下方插入', children: below },
    ...(props.latexState ? [documentLatexMountedMenuItem(control.empty ? 'insert-quick-bottom' : 'edit-quick-bottom', props.latexState, t)] : [])]
}
function itemsFor(kind: MenuKind, control?: Control): RibbonItem[] {
  if (kind === 'structure') return props.active ? documentStructureContextItems(props.active) : []
  if (!control) return []
  return kind === 'paragraph' ? paragraphItems(control) : documentStructuredMountedMenuItems(control.context, kind === 'table' ? 'table-below' : 'content-insert')
}
function convertItem(item: RibbonItem): MenuItemType {
  return { key: item.id, label: item.title, disabled: item.disabled,
    ...(item.children?.length ? { children: item.children.filter(child => !child.hidden).map(convertItem) } : {}) }
}
function findItem(items: readonly RibbonItem[], key: string): RibbonItem | undefined {
  for (const item of items) { if (item.id === key) return item; const child = item.children && findItem(item.children, key); if (child) return child }
  return undefined
}
function selectedKeys(items: readonly RibbonItem[]): string[] {
  return items.flatMap(item => [...(item.activated ? [item.id] : []), ...selectedKeys(item.children ?? [])])
}
function menuConfig(kind: MenuKind, control?: Control) {
  const items = menu.value?.kind === kind && menu.value.key === control?.key ? menu.value.items : itemsFor(kind, control)
  return { items: items.filter(item => !item.hidden).map(convertItem), selectedKeys: selectedKeys(items), onClick: clickMenu }
}
function closeMenu(refocus = false) {
  const pending = menu.value; menu.value = undefined
  if (refocus && pending && props.targetCurrent(pending.target)) emit('focus', pending.target)
}
function closeCode(refocus = false) {
  const pending = code.value; code.value = undefined
  if (refocus && pending && props.targetCurrent(pending.target)) emit('focus', pending.target)
}
function openMenu(kind: MenuKind, open: boolean, control?: Control) {
  if (!open) { if (menu.value?.kind === kind && menu.value.key === control?.key) closeMenu(true); return }
  if (props.disabled || control?.context.disabled) return
  const target = props.captureTarget(), resolved = control ?? (kind === 'paragraph' ? props.paragraph : undefined)
  if (!props.targetCurrent(target)) return
  closeCode(); menu.value = { kind, key: resolved?.key, target, items: itemsFor(kind, resolved),
    title: kind === 'table' ? '在表格下方插入' : kind === 'content' ? '插入内容' : kind === 'structure' ? '高级块操作' : '段落操作' }
}
function openCode(kind: CodeControlKind, open: boolean, target = props.captureTarget()) {
  if (!open) { if (code.value?.kind === kind) closeCode(true); return }
  const active = props.active
  if (props.disabled || !active?.code || !props.targetCurrent(target)) return
  closeMenu(); code.value = { kind, blockId: active.blockId, target, config: { ...active.code }, fontFamily: active.fontFamily ?? 'monospace',
    actions: documentStructureContextItems(active).filter(item => ['doc-context-structured-properties', 'doc-context-structured-unwrap', 'doc-context-structured-exit', 'doc-context-structured-remove'].includes(item.id)) }
}
function dispatch(item: RibbonItem, target: DocumentStructuredCanvasTarget) {
  if (item.hidden || item.disabled || !item.commandId || !props.targetCurrent(target)) return
  const command = documentStructureCommandFromMenu(item.commandId, item.params)
  if (command) emit('command', { ...target, command })
  else if (item.commandId === 'doc.operation.edit-code-font' || item.commandId === 'doc.operation.edit-code-settings') openCode(item.commandId === 'doc.operation.edit-code-font' ? 'font' : 'settings', true, target)
  else emit('operation', { ...target, id: item.commandId, params: item.params })
}
function clickMenu(info: { key: string | number }) {
  const pending = menu.value, item = pending && findItem(pending.items, String(info.key))
  closeMenu()
  if (pending && item) dispatch(item, pending.target)
}
function applyCode(value: { patch?: Partial<DocumentCodeConfig>; fontFamily?: string }) {
  const pending = code.value; closeCode()
  if (!pending || !props.targetCurrent(pending.target)) return
  const command = value.fontFamily === undefined
    ? documentStructureCommandFromMenu('doc.command.update-code-settings', { blockId: pending.blockId, patch: value.patch })
    : documentStructureCommandFromMenu('doc.command.set-code-font-family', { blockId: pending.blockId, fontFamily: value.fontFamily })
  if (command) emit('command', { ...pending.target, command })
}
function codeAction(info: { key: string | number }) {
  const pending = code.value, item = pending && findItem(pending.actions, String(info.key)); closeCode()
  if (pending && item) dispatch(item, pending.target)
}
function prepareProperties() { if (!props.disabled && props.active) direct.value = { target: props.captureTarget(), blockId: props.active.blockId } }
function properties() {
  const pending = direct.value ?? (props.active ? { target: props.captureTarget(), blockId: props.active.blockId } : undefined)
  direct.value = undefined
  if (pending && props.targetCurrent(pending.target)) emit('operation', { ...pending.target, id: 'doc.operation.edit-structured-block', params: { blockId: pending.blockId } })
}
function prepareLatexItem(item: RibbonItem, event: PointerEvent) {
  event.preventDefault()
  if (props.disabled || item.disabled) return
  const target = props.captureTarget()
  if (props.targetCurrent(target)) directLatex.value = { target, itemId: item.id }
}
function clickLatexItem(item: RibbonItem) {
  const pending = directLatex.value; directLatex.value = undefined
  dispatch(item, pending?.itemId === item.id ? pending.target : props.captureTarget())
}
function cancelInteraction(refocus = false) { closeMenu(refocus); closeCode(refocus); direct.value = undefined; directLatex.value = undefined }
function hasInteraction() { return !!menu.value || !!code.value }
defineExpose({ cancelInteraction, hasInteraction })
</script>

<template>
  <div class="structured-quick-tools" :class="{ 'structured-quick-tools--mobile': mobile }" :data-document-mobile-quick-tools="mobile || undefined">
  <div v-if="floatLatex" class="structured-latex-toolbar" :style="mobile ? undefined : floatLatex.style" role="toolbar" :aria-label="t('document.latex.floating-tools')" data-document-float-text-style>
    <Tooltip :title="floatLatex.item.disabledReason ?? floatLatex.item.title">
      <Button size="small" :disabled="disabled || floatLatex.item.disabled" :aria-label="floatLatex.item.title" data-document-latex-mount="float-text-style"
        @pointerdown="event => floatLatex && prepareLatexItem(floatLatex.item, event)" @click="clickLatexItem(floatLatex.item)"><Sigma :size="14" />{{ floatLatex.item.title }}</Button>
    </Tooltip>
  </div>
  <div v-if="mobile && mobileLatex" class="structured-latex-toolbar" role="toolbar" :aria-label="t('document.latex.floating-tools')" data-document-mobile-latex-actions>
    <Tooltip v-for="item in mobileLatex.items" :key="item.id" :title="item.disabledReason ?? item.title">
      <Button size="small" :danger="item.commandId === 'doc.command.remove-latex-formula'" :disabled="disabled || item.disabled" :aria-label="item.title"
        :data-document-latex-action="item.id" @pointerdown="event => prepareLatexItem(item, event)" @click="clickLatexItem(item)">{{ item.title }}</Button>
    </Tooltip>
  </div>
  <div v-if="active" class="structured-toolbar" :class="{ 'structured-toolbar--mobile': mobile }" :style="floatingStyle" data-document-structured-toolbar :data-document-mobile-toolbar="mobile || undefined">
    <template v-if="active.code">
      <Popover v-if="!mobile" :open="code?.kind === 'font'" trigger="click" placement="bottomRight" destroy-on-hidden @open-change="open => openCode('font', open)">
        <template #content><DocumentCodeControls v-if="code?.kind === 'font'" kind="font" :config="code.config" :font-family="code.fontFamily" :disabled="disabled" @apply="applyCode" @cancel="closeCode(true)" /></template>
        <Button size="small" :disabled="disabled" aria-label="代码块字体">{{ active.fontFamily ?? '等宽字体' }}</Button>
      </Popover>
      <Button v-else size="small" :disabled="disabled" aria-label="代码块字体" @click="openCode('font', true)">字体</Button>
      <Popover v-if="!mobile" :open="code?.kind === 'language'" trigger="click" placement="bottomRight" destroy-on-hidden @open-change="open => openCode('language', open)">
        <template #content><DocumentCodeControls v-if="code?.kind === 'language'" kind="language" :config="code.config" :font-family="code.fontFamily" :disabled="disabled" @apply="applyCode" @cancel="closeCode(true)" /></template>
        <Button size="small" :disabled="disabled" aria-label="代码块语言选择">{{ languageTitle }}</Button>
      </Popover>
      <Button v-else size="small" :disabled="disabled" aria-label="代码块语言选择" @click="openCode('language', true)">语言</Button>
      <Popover v-if="!mobile" :open="code?.kind === 'settings'" trigger="click" placement="bottomRight" destroy-on-hidden @open-change="open => openCode('settings', open)">
        <template #content>
          <template v-if="code?.kind === 'settings'">
            <DocumentCodeControls kind="settings" :config="code.config" :font-family="code.fontFamily" :disabled="disabled" @apply="applyCode" @cancel="closeCode(true)" />
            <Divider /><Menu :items="code.actions.map(convertItem)" :selectable="false" @click="codeAction" />
          </template>
        </template>
        <Button size="small" :disabled="disabled" aria-label="高级块操作">更多</Button>
      </Popover>
      <Button v-else size="small" :disabled="disabled" aria-label="高级块操作" @click="openCode('settings', true)">更多</Button>
    </template>
    <template v-else>
      <Tooltip :title="`${active.title}属性`"><Button size="small" :disabled="disabled" @pointerdown="prepareProperties" @click="properties">属性</Button></Tooltip>
      <Dropdown v-if="!mobile" :open="menu?.kind === 'structure'" :trigger="['click']" :menu="menuConfig('structure')" :disabled="disabled" destroy-on-hidden @open-change="open => openMenu('structure', open)"><Button size="small" :disabled="disabled" aria-label="高级块操作">更多</Button></Dropdown>
      <Button v-else size="small" :disabled="disabled" aria-label="高级块操作" @click="openMenu('structure', true)">更多</Button>
    </template>
  </div>
  <div v-if="paragraph" class="structured-paragraph-control" :style="paragraphStyle" data-document-paragraph-insert :data-document-paragraph-menu="paragraph.empty ? 'quick-insert' : 'quick-edit'">
    <Dropdown v-if="!mobile" :open="menu?.kind === 'paragraph'" :trigger="['click']" :menu="menuConfig('paragraph', paragraph)" :disabled="disabled" destroy-on-hidden @open-change="open => openMenu('paragraph', open, paragraph)">
      <Button type="text" size="small" :disabled="disabled" :aria-label="paragraph.empty ? '空段落插入' : '段落操作'">{{ paragraph.empty ? '＋' : 'T' }}</Button>
    </Dropdown>
    <Button v-else size="small" :disabled="disabled" aria-label="段落操作" @click="openMenu('paragraph', true, paragraph)">段落</Button>
  </div>
  </div>
  <div v-if="contentInsert" class="structured-content-insert" :style="contentInsert.style" data-document-content-insert>
    <Dropdown v-if="!mobile" :open="menu?.kind === 'content'" :trigger="['click']" :menu="menuConfig('content', contentInsert)" :disabled="disabled" destroy-on-hidden @open-change="open => openMenu('content', open, contentInsert)"><Button type="text" size="small" :disabled="disabled" aria-label="在下方插入内容">＋</Button></Dropdown>
    <Button v-else size="small" :disabled="disabled" aria-label="在下方插入内容" @click="openMenu('content', true, contentInsert)">＋</Button>
  </div>
  <div v-for="control in tables" :key="control.key" class="structured-table-insert" :style="control.style" data-document-table-below-insert :data-document-table-id="control.context.afterBlockId">
    <Dropdown v-if="!mobile" :open="menu?.kind === 'table' && menu.key === control.key" :trigger="['click']" :menu="menuConfig('table', control)" :disabled="disabled" destroy-on-hidden @open-change="open => openMenu('table', open, control)"><Button size="small" :disabled="disabled" aria-label="在表格下方插入">＋</Button></Dropdown>
    <Button v-else size="small" :disabled="disabled" aria-label="在表格下方插入" @click="openMenu('table', true, control)">＋</Button>
  </div>
  <Drawer :open="mobile && !!menu" :title="menu?.title" placement="bottom" :size="360" :focusable="{ focusTriggerAfterClose: false }" destroy-on-hidden @close="closeMenu(true)">
    <Menu v-if="menu" :items="menu.items.filter(item => !item.hidden).map(convertItem)" :selected-keys="selectedKeys(menu.items)" mode="inline" @click="clickMenu" />
  </Drawer>
  <Drawer :open="mobile && !!code" :title="codeTitle" placement="bottom" :size="440" :focusable="{ focusTriggerAfterClose: false }" destroy-on-hidden @close="closeCode(true)">
    <template v-if="code">
      <DocumentCodeControls :kind="code.kind" :config="code.config" :font-family="code.fontFamily" :disabled="disabled" @apply="applyCode" @cancel="closeCode(true)" />
      <template v-if="code.kind === 'settings'"><Divider /><Menu :items="code.actions.map(convertItem)" :selectable="false" @click="codeAction" /></template>
    </template>
  </Drawer>
</template>

<style scoped>
.structured-quick-tools { display: contents; }
.structured-quick-tools--mobile { position: absolute; left: 8px; bottom: 8px; display: flex; flex-direction: column; align-items: flex-start; gap: 6px; max-width: calc(100% - 16px); }
.structured-quick-tools--mobile > .structured-latex-toolbar, .structured-quick-tools--mobile > .structured-toolbar, .structured-quick-tools--mobile > .structured-paragraph-control { position: static; max-width: 100%; }
.structured-toolbar { position: absolute; display: flex; gap: 4px; padding: 3px; border: 1px solid var(--structured-border); border-radius: 6px; background: var(--structured-bg); box-shadow: var(--structured-shadow); pointer-events: auto; }
.structured-toolbar--mobile { flex-wrap: wrap; }
.structured-latex-toolbar { position: absolute; display: flex; gap: 4px; padding: 4px; border: 1px solid var(--structured-border); border-radius: 6px; background: var(--structured-bg); box-shadow: var(--structured-shadow); pointer-events: auto; }
.structured-latex-toolbar :deep(.ant-btn) { display: inline-flex; align-items: center; gap: 4px; }
.structured-paragraph-control, .structured-content-insert, .structured-table-insert { position: absolute; pointer-events: auto; }
.structured-paragraph-control :deep(.ant-btn), .structured-content-insert :deep(.ant-btn) { padding-inline: 5px; }
</style>
