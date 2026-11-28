<script setup lang="ts">
import { Button, Tooltip, theme } from 'antdv-next'
import { ListTree, PanelLeftClose } from 'lucide-vue-next'
import { computed, inject, nextTick, onActivated, onBeforeUnmount, onDeactivated, onMounted, ref, useId, watch } from 'vue'
import type { DocumentModel } from '../../domain/model'
import { buildDocumentOutline, nextOutlineFocusId, type DocumentOutlineItem } from '../../features/document/services/document-outline'
import { OfficeTranslatorKey } from '../../i18n/office-i18n'

const props = withDefaults(defineProps<{
  model: DocumentModel
  activeId?: string
  mode?: 'side-bar' | 'float'
  open?: boolean
  maxHeight?: number
  maxWidth?: number
  disabled?: boolean
}>(), {
  activeId: undefined,
  mode: 'side-bar',
  open: undefined,
  maxHeight: undefined,
  maxWidth: 256,
  disabled: false,
})

const emit = defineEmits<{
  navigate: [item: DocumentOutlineItem]
  'update:open': [open: boolean]
}>()

const panelId = useId()
const { token } = theme.useToken()
const translate = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const title = computed(() => translate('docs-ui.outline.title', '文档大纲'))
const root = ref<HTMLElement | null>(null)
const scrollContainer = ref<HTMLElement | null>(null)
const internalOpen = ref(props.mode === 'side-bar')
let active = true
let listenerDocument: Document | undefined
const focusedId = ref<string>()
const items = computed(() => buildDocumentOutline(props.model))
const expanded = computed(() => props.open ?? internalOpen.value)
const headingCount = computed(() => items.value.filter((item) => !item.isTitle).length)
const panelWidth = computed(() => Math.min(256, Math.max(32, Number.isFinite(props.maxWidth) ? props.maxWidth : 256)))
const panelStyle = computed(() => ({
  width: `${expanded.value ? panelWidth.value : 40}px`,
  maxHeight: props.maxHeight != null && Number.isFinite(props.maxHeight) ? `${Math.max(32, props.maxHeight)}px` : undefined,
  '--outline-border': token.value.colorBorderSecondary, '--outline-background': token.value.colorBgContainer,
  '--outline-text': token.value.colorTextSecondary, '--outline-title': token.value.colorText,
  '--outline-hover': token.value.colorFillSecondary, '--outline-primary': token.value.colorPrimary,
  '--outline-active': token.value.colorPrimaryBg, '--outline-shadow': token.value.boxShadowSecondary,
  '--outline-disabled': token.value.colorTextDisabled,
}))
const tabStopId = computed(() => items.value.find((item) => item.id === focusedId.value)?.id
  ?? items.value.find((item) => item.id === props.activeId)?.id
  ?? items.value[0]?.id)

function setOpen(value: boolean, restoreFocus = false): void {
  if (props.open === undefined) internalOpen.value = value
  emit('update:open', value)
  const source = props.model
  if (restoreFocus) void nextTick(() => { if (active && props.model === source) toggleElement()?.focus() })
}

function toggleElement(): HTMLButtonElement | null {
  return root.value?.querySelector<HTMLButtonElement>('[data-outline-toggle]') ?? null
}

function itemElement(id: string): HTMLButtonElement | undefined {
  return Array.from(root.value?.querySelectorAll<HTMLButtonElement>('[data-outline-item]') ?? [])
    .find((element) => element.dataset.outlineItem === id)
}

function revealActive(): void {
  const container = scrollContainer.value
  const element = props.activeId ? itemElement(props.activeId) : undefined
  if (!container || !element || !expanded.value) return
  const viewport = container.getBoundingClientRect()
  const item = element.getBoundingClientRect()
  if (item.top < viewport.top) container.scrollTop -= viewport.top - item.top
  else if (item.bottom > viewport.bottom) container.scrollTop += item.bottom - viewport.bottom
}

function navigate(item: DocumentOutlineItem): void {
  if (!active || props.disabled) return
  focusedId.value = item.id
  emit('navigate', item)
  if (props.mode === 'float') setOpen(false, true)
}

function onKeydown(event: KeyboardEvent): void {
  if (event.key === 'Escape' && expanded.value) {
    event.preventDefault()
    event.stopPropagation()
    setOpen(false, true)
    return
  }
  if (event.altKey || event.ctrlKey || event.metaKey) return
  const target = event.target
  const item = Array.from(root.value?.querySelectorAll<HTMLElement>('[data-outline-item]') ?? []).find(element => element === target)
  if (!item?.dataset.outlineItem) return
  const nextId = nextOutlineFocusId(items.value, item.dataset.outlineItem, event.key)
  if (!nextId) return
  event.preventDefault()
  event.stopPropagation()
  focusedId.value = nextId
  itemElement(nextId)?.focus()
}

function onOutsidePointer(event: PointerEvent): void {
  const element = root.value
  if (active && element && props.mode === 'float' && expanded.value && !event.composedPath().includes(element)) setOpen(false)
}

function unbindOutsidePointer(): void {
  listenerDocument?.removeEventListener('pointerdown', onOutsidePointer)
  listenerDocument = undefined
}

function bindOutsidePointer(): void {
  active = true
  unbindOutsidePointer()
  listenerDocument = root.value?.ownerDocument
  listenerDocument?.addEventListener('pointerdown', onOutsidePointer)
}

watch(() => props.mode, (mode) => { if (props.open === undefined) internalOpen.value = mode === 'side-bar' })
watch([() => props.activeId, expanded, items], () => { void nextTick(revealActive) }, { flush: 'post' })
watch(items, (value) => { if (!value.some((item) => item.id === focusedId.value)) focusedId.value = undefined })
watch(() => props.model.id, () => { focusedId.value = undefined; if (props.open === undefined) internalOpen.value = props.mode === 'side-bar' })
onMounted(bindOutsidePointer)
onActivated(bindOutsidePointer)
onDeactivated(() => { active = false; unbindOutsidePointer() })
onBeforeUnmount(() => { active = false; unbindOutsidePointer() })

defineExpose({ revealActive })
</script>

<template>
  <aside ref="root" class="document-outline" :class="[`document-outline--${mode}`, { 'is-open': expanded }]" :style="panelStyle" @keydown="onKeydown">
    <div class="document-outline__header">
      <Tooltip :title="`${translate(expanded ? 'document.heading.collapse' : 'document.heading.expand', expanded ? '收起' : '展开')} ${title}`" :trigger="['hover', 'focus']">
        <Button
          data-outline-toggle
          class="document-outline__toggle"
          type="text"
          :aria-label="title"
          :aria-expanded="expanded"
          :aria-controls="expanded ? panelId : undefined"
          @click="setOpen(!expanded)"
        >
          <PanelLeftClose v-if="expanded" :size="18" aria-hidden="true" />
          <ListTree v-else :size="18" aria-hidden="true" />
        </Button>
      </Tooltip>
      <span v-if="expanded" class="document-outline__label">{{ title }}</span>
    </div>
    <nav v-if="expanded" :id="panelId" ref="scrollContainer" class="document-outline__navigation" :aria-label="title">
      <ol class="document-outline__list">
        <li v-for="item in items" :key="item.id" :class="{ 'document-outline__title': item.isTitle }">
          <Tooltip :title="item.text" :trigger="['hover', 'focus']">
            <Button
              class="document-outline__item"
              :class="{ 'is-active': activeId === item.id, 'is-title': item.isTitle }"
              type="text"
              :data-outline-item="item.id"
              :aria-current="activeId === item.id ? 'location' : undefined"
              :disabled="disabled"
              :tabindex="tabStopId === item.id ? 0 : -1"
              :style="{ paddingInlineStart: `${item.isTitle ? 4 : 12 + item.depth * 12}px` }"
              @focus="focusedId = item.id"
              @click="navigate(item)"
            >
              <span>{{ item.text }}</span>
            </Button>
          </Tooltip>
        </li>
      </ol>
      <p v-if="!headingCount" class="document-outline__empty">文档中暂无标题。为段落应用标题样式后，将在这里显示。</p>
    </nav>
  </aside>
</template>

<style scoped>
.document-outline { position: relative; display: flex; flex: 0 0 auto; flex-direction: column; min-width: 0; min-height: 0; max-width: 100%; overflow: hidden; box-sizing: border-box; padding: 12px 8px; border-inline-end: 1px solid var(--outline-border); background: var(--outline-background); color: var(--outline-text); }
.document-outline__header { display: flex; flex: 0 0 32px; align-items: center; gap: 4px; min-width: 0; }
.document-outline__label { overflow: hidden; color: var(--outline-text); font-size: 12px; font-weight: 500; white-space: nowrap; text-overflow: ellipsis; }
.document-outline__toggle.ant-btn { display: inline-flex; flex: 0 0 32px; align-items: center; justify-content: center; width: 32px; height: 32px; padding: 0; color: var(--outline-text); }
.document-outline__navigation { min-height: 0; overflow: auto; overscroll-behavior: contain; scrollbar-width: thin; }
.document-outline__list { margin: 0; padding: 0 0 8px; list-style: none; }
.document-outline__title { margin-bottom: 8px; }
.document-outline__item.ant-btn { display: flex; align-items: center; justify-content: flex-start; width: 100%; height: auto; min-height: 32px; padding-block: 6px; padding-inline-end: 8px; border-radius: 4px; color: var(--outline-text); font-size: 13px; line-height: 20px; text-align: start; }
.document-outline__item span { display: block; min-width: 0; overflow: hidden; white-space: nowrap; text-overflow: ellipsis; }
.document-outline__item.is-title { color: var(--outline-title); font-weight: 500; }
.document-outline__item.is-title span { display: -webkit-box; -webkit-box-orient: vertical; -webkit-line-clamp: 2; white-space: normal; overflow-wrap: anywhere; }
.document-outline__item.ant-btn:not(:disabled):hover { color: var(--outline-title); background: var(--outline-hover); }
.document-outline__item.is-active.ant-btn { color: var(--outline-primary); background: var(--outline-active); }
.document-outline__item.ant-btn:disabled { color: var(--outline-disabled); background: transparent; }
.document-outline__item.ant-btn:focus-visible, .document-outline__toggle.ant-btn:focus-visible { outline: 2px solid var(--outline-primary); outline-offset: -2px; }
.document-outline__empty { margin: 4px 8px 12px; color: var(--outline-text); font-size: 12px; line-height: 1.8; }
.document-outline--float { position: absolute; inset-inline-start: 16px; top: 16px; z-index: 30; max-height: calc(100% - 32px); padding: 4px; border: 0; border-radius: 6px; background: transparent; }
.document-outline--float.is-open { background: var(--outline-background); box-shadow: var(--outline-shadow); }
@media print { .document-outline { display: none !important; } }
</style>
