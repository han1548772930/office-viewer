<script setup lang="ts">
import { Button, Tooltip } from 'antdv-next'
import { ListTree, PanelLeftClose } from 'lucide-vue-next'
import { computed, nextTick, onBeforeUnmount, onMounted, ref, useId, watch } from 'vue'
import type { DocumentModel } from '../../domain/model'
import { buildDocumentOutline, nextOutlineFocusId, type DocumentOutlineItem } from '../../features/document/services/document-outline'

const props = withDefaults(defineProps<{
  model: DocumentModel
  activeId?: string
  mode?: 'side-bar' | 'float'
  open?: boolean
  maxHeight?: number
  maxWidth?: number
}>(), {
  activeId: undefined,
  mode: 'side-bar',
  open: undefined,
  maxHeight: undefined,
  maxWidth: 256,
})

const emit = defineEmits<{
  navigate: [item: DocumentOutlineItem]
  'update:open': [open: boolean]
}>()

const panelId = useId()
const root = ref<HTMLElement | null>(null)
const scrollContainer = ref<HTMLElement | null>(null)
const internalOpen = ref(props.mode === 'side-bar')
const focusedId = ref<string>()
const items = computed(() => buildDocumentOutline(props.model))
const expanded = computed(() => props.open ?? internalOpen.value)
const headingCount = computed(() => items.value.filter((item) => !item.isTitle).length)
const panelWidth = computed(() => Math.min(256, Math.max(32, Number.isFinite(props.maxWidth) ? props.maxWidth : 256)))
const panelStyle = computed(() => ({
  width: `${expanded.value ? panelWidth.value : 40}px`,
  maxHeight: props.maxHeight != null && Number.isFinite(props.maxHeight) ? `${Math.max(32, props.maxHeight)}px` : undefined,
}))
const tabStopId = computed(() => items.value.find((item) => item.id === focusedId.value)?.id
  ?? items.value.find((item) => item.id === props.activeId)?.id
  ?? items.value[0]?.id)

function setOpen(value: boolean, restoreFocus = false): void {
  if (props.open === undefined) internalOpen.value = value
  emit('update:open', value)
  if (restoreFocus) void nextTick(() => toggleElement()?.focus())
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
  if (!(target instanceof HTMLElement) || !target.dataset.outlineItem) return
  const nextId = nextOutlineFocusId(items.value, target.dataset.outlineItem, event.key)
  if (!nextId) return
  event.preventDefault()
  event.stopPropagation()
  focusedId.value = nextId
  itemElement(nextId)?.focus()
}

function onOutsidePointer(event: PointerEvent): void {
  if (props.mode === 'float' && expanded.value && event.target instanceof Node && !root.value?.contains(event.target)) setOpen(false)
}

watch(() => props.mode, (mode) => { if (props.open === undefined) internalOpen.value = mode === 'side-bar' })
watch([() => props.activeId, expanded, items], () => { void nextTick(revealActive) }, { flush: 'post' })
watch(items, (value) => { if (!value.some((item) => item.id === focusedId.value)) focusedId.value = undefined })
onMounted(() => document.addEventListener('pointerdown', onOutsidePointer))
onBeforeUnmount(() => document.removeEventListener('pointerdown', onOutsidePointer))

defineExpose({ revealActive })
</script>

<template>
  <aside ref="root" class="document-outline" :class="[`document-outline--${mode}`, { 'is-open': expanded }]" :style="panelStyle" @keydown="onKeydown">
    <div class="document-outline__header">
      <Tooltip :title="expanded ? '收起文档大纲' : '展开文档大纲'">
        <Button
          data-outline-toggle
          class="document-outline__toggle"
          type="text"
          aria-label="文档大纲"
          :aria-expanded="expanded"
          :aria-controls="expanded ? panelId : undefined"
          @click="setOpen(!expanded)"
        >
          <PanelLeftClose v-if="expanded" :size="18" aria-hidden="true" />
          <ListTree v-else :size="18" aria-hidden="true" />
        </Button>
      </Tooltip>
      <span v-if="expanded" class="document-outline__label">文档大纲</span>
    </div>
    <nav v-if="expanded" :id="panelId" ref="scrollContainer" class="document-outline__navigation" aria-label="文档大纲">
      <ol class="document-outline__list">
        <li v-for="item in items" :key="item.id" :class="{ 'document-outline__title': item.isTitle }">
          <Button
            class="document-outline__item"
            :class="{ 'is-active': activeId === item.id, 'is-title': item.isTitle }"
            type="text"
            :data-outline-item="item.id"
            :aria-current="activeId === item.id ? 'location' : undefined"
            :title="item.text"
            :tabindex="tabStopId === item.id ? 0 : -1"
            :style="{ paddingInlineStart: `${item.isTitle ? 4 : 12 + item.depth * 12}px` }"
            @focus="focusedId = item.id"
            @click="navigate(item)"
          >
            <span>{{ item.text }}</span>
          </Button>
        </li>
      </ol>
      <p v-if="!headingCount" class="document-outline__empty">文档中暂无标题。为段落应用标题样式后，将在这里显示。</p>
    </nav>
  </aside>
</template>

<style scoped>
.document-outline { position: relative; display: flex; flex: 0 0 auto; flex-direction: column; min-width: 0; min-height: 0; max-width: 100%; overflow: hidden; box-sizing: border-box; padding: 12px 8px; border-inline-end: 1px solid #e8ebef; background: #f8f9fb; color: #596371; }
.document-outline__header { display: flex; flex: 0 0 32px; align-items: center; gap: 4px; min-width: 0; }
.document-outline__label { overflow: hidden; color: #616d7c; font-size: 12px; font-weight: 500; white-space: nowrap; text-overflow: ellipsis; }
.document-outline__toggle.ant-btn { display: inline-flex; flex: 0 0 32px; align-items: center; justify-content: center; width: 32px; height: 32px; padding: 0; color: #687484; }
.document-outline__navigation { min-height: 0; overflow: auto; overscroll-behavior: contain; scrollbar-width: thin; }
.document-outline__list { margin: 0; padding: 0 0 8px; list-style: none; }
.document-outline__title { margin-bottom: 8px; }
.document-outline__item.ant-btn { display: flex; align-items: center; justify-content: flex-start; width: 100%; height: auto; min-height: 32px; padding-block: 6px; padding-inline-end: 8px; border-radius: 4px; color: #687484; font-size: 13px; line-height: 20px; text-align: start; }
.document-outline__item span { display: block; min-width: 0; overflow: hidden; white-space: nowrap; text-overflow: ellipsis; }
.document-outline__item.is-title { color: #424e5f; font-weight: 500; }
.document-outline__item.is-title span { display: -webkit-box; -webkit-box-orient: vertical; -webkit-line-clamp: 2; white-space: normal; overflow-wrap: anywhere; }
.document-outline__item.ant-btn:hover { color: #23364f; background: #edf0f5; }
.document-outline__item.is-active.ant-btn { color: #2563ba; background: #e9f1ff; }
.document-outline__item.ant-btn:focus-visible, .document-outline__toggle.ant-btn:focus-visible { outline: 2px solid #436fea; outline-offset: -2px; }
.document-outline__empty { margin: 4px 8px 12px; color: #7b8797; font-size: 12px; line-height: 1.8; }
.document-outline--float { position: absolute; inset-inline-start: 16px; top: 16px; z-index: 30; max-height: calc(100% - 32px); padding: 4px; border: 0; border-radius: 6px; background: transparent; }
.document-outline--float.is-open { background: #fff; box-shadow: 0 4px 20px #23334d26; }
@media print { .document-outline { display: none !important; } }
</style>
