<script setup lang="ts">
import { computed, h, inject, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import { Button, Menu } from 'antdv-next'
import type { ItemType } from 'antdv-next/es/menu/interface'
import { ArrowDownToLine, ArrowLeftToLine, ArrowRightToLine, ArrowUpToLine, ClipboardPaste, Columns3, Copy, Eraser, Eye, EyeOff, Hash, LockKeyhole, Minus, PinOff, Plus, Rows3, Scissors, ShieldCheck, Snowflake, Trash2 } from 'lucide-vue-next'
import type { GridHit } from '../../rendering/spreadsheet/layout.ts'
import { createSheetContextMenuItems, type SheetContextMenuItem } from './sheet-context-menu.ts'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = defineProps<{
  hit: GridHit | null
  sheetId: string
  x: number
  y: number
  open: boolean
  canDeleteRow: boolean
  canDeleteColumn: boolean
  frozen: boolean
  hasHiddenRows: boolean
  hasHiddenColumns: boolean
  canConvertText?: boolean
  canShiftHorizontally?: boolean
  canShiftVertically?: boolean
}>()
const emit = defineEmits<{
  'update:open': [value: boolean]
  command: [commandId: string, params?: unknown]
}>()

const panel = ref<HTMLElement>()
const translate = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
function translateItems(source: SheetContextMenuItem[]): SheetContextMenuItem[] {
  return source.flatMap(item => [
    ...(item.separatorBefore ? [{ id: `${item.id}-separator`, label: '', commandId: '', type: 'divider' as const }] : []),
    { ...item, separatorBefore: false },
  ].map(entry => ({
    ...entry,
    label: entry.type === 'divider' ? '' : translate(`sheet.context-menu.${entry.id}`, entry.label),
    children: entry.children ? translateItems(entry.children) : undefined,
  })))
}
const items = computed(() => props.hit ? translateItems(createSheetContextMenuItems(props.hit, props)) : [])
const itemById = computed(() => {
  const entries = new Map<string, (typeof items.value)[number]>()
  const visit = (entry: (typeof items.value)[number]) => {
    entries.set(entry.id, entry)
    entry.children?.forEach(visit)
  }
  items.value.forEach(visit)
  return entries
})
const quickItems = computed(() => items.value.filter(item => item.quick))
const regularItems = computed(() => items.value.filter(item => !item.quick))
const menuItems = computed<ItemType[]>(() => toMenuItems(regularItems.value))
const position = computed(() => ({ left: `${Math.max(4, props.x)}px`, top: `${Math.max(4, props.y)}px` }))
const menuRef = ref<InstanceType<typeof Menu> | null>(null)

const menuIcons = {
  CopyDoubleIcon: Copy,
  CutIcon: Scissors,
  PasteSpecialDoubleIcon: ClipboardPaste,
  ClearFormatDoubleIcon: Eraser,
  FreezeToSelectedIcon: Snowflake,
  FreezeRowIcon: Rows3,
  FreezeColumnIcon: Columns3,
  CancelFreezeIcon: PinOff,
  ProtectIcon: ShieldCheck,
  LockIcon: LockKeyhole,
  EyeIcon: Eye,
  HideDoubleIcon: EyeOff,
  InsertDoubleIcon: Plus,
  ReduceDoubleIcon: Minus,
  InsertRowAboveDoubleIcon: ArrowUpToLine,
  InsertRowBelowDoubleIcon: ArrowDownToLine,
  LeftInsertColumnDoubleIcon: ArrowLeftToLine,
  RightInsertColumnDoubleIcon: ArrowRightToLine,
  InsertCellShiftRightDoubleIcon: ArrowRightToLine,
  InsertCellDownDoubleIcon: ArrowDownToLine,
  DeleteCellShiftLeftDoubleIcon: ArrowLeftToLine,
  DeleteCellShiftUpDoubleIcon: ArrowUpToLine,
  DeleteRowDoubleIcon: Trash2,
  DeleteColumnDoubleIcon: Trash2,
  AdjustHeightDoubleIcon: Rows3,
  AdjustWidthDoubleIcon: Columns3,
  AutoHeightDoubleIcon: Rows3,
  AutoWidthDoubleIcon: Columns3,
  ConvertToNumberIcon: Hash,
} as const

function toMenuItem(item: (typeof items.value)[number]): ItemType {
  const Icon = iconComponent(item)
  return {
    key: item.id,
    label: item.label,
    title: item.label,
    disabled: item.disabled,
    icon: Icon ? h(Icon, { size: 16 }) : undefined,
    children: item.children ? toMenuItems(item.children) : undefined,
    ...(item.children ? { popupClassName: 'sheet-context-menu__submenu' } : {}),
  }
}
function iconComponent(item: (typeof items.value)[number]) {
  return item.icon ? menuIcons[item.icon as keyof typeof menuIcons] : undefined
}
function toMenuItems(source: typeof items.value): ItemType[] {
  return source.map(item => item.type === 'divider'
    ? { type: 'divider' as const, key: item.id }
    : toMenuItem(item))
}

function close() { emit('update:open', false) }
function execute(item: (typeof items.value)[number] | undefined) {
  if (!item || item.children?.length) return
  if (item.disabled) return
  close()
  emit('command', item.commandId, item.params)
}
function onKeydown(event: KeyboardEvent) {
  if (props.open && event.key === 'Escape') { event.preventDefault(); close() }
}
function onOutsidePointer(event: PointerEvent) {
  const target = event.target instanceof Element ? event.target : null
  if (props.open && !panel.value?.contains(event.target as Node) && !target?.closest('.sheet-context-menu__submenu')) close()
}
function onViewportChange() { if (props.open) close() }

watch(() => [props.open, props.x, props.y] as const, async ([open]) => {
  if (!open) return
  await nextTick()
  const bounds = panel.value?.getBoundingClientRect()
  if (bounds && (bounds.right > window.innerWidth - 4 || bounds.bottom > window.innerHeight - 4)) {
    panel.value!.style.left = `${Math.max(4, Math.min(props.x, window.innerWidth - bounds.width - 4))}px`
    panel.value!.style.top = `${Math.max(4, Math.min(props.y, window.innerHeight - bounds.height - 4))}px`
  }
  menuRef.value?.focus()
})
onMounted(() => {
  document.addEventListener('pointerdown', onOutsidePointer)
  document.addEventListener('keydown', onKeydown)
  window.addEventListener('resize', onViewportChange)
  window.addEventListener('scroll', onViewportChange, true)
})
onBeforeUnmount(() => {
  document.removeEventListener('pointerdown', onOutsidePointer)
  document.removeEventListener('keydown', onKeydown)
  window.removeEventListener('resize', onViewportChange)
  window.removeEventListener('scroll', onViewportChange, true)
})
</script>

<template>
  <Teleport to="body">
    <div v-if="open && items.length" ref="panel" class="sheet-context-menu" :style="position">
      <div v-if="quickItems.length" class="sheet-context-menu__quick" role="group" aria-label="快捷操作">
        <Button
          v-for="item in quickItems"
          :key="item.id"
          type="text"
          size="small"
          class="sheet-context-menu__quick-item"
          :title="item.label"
          :aria-label="item.label"
          :disabled="item.disabled"
          @click="execute(item)"
        >
          <component :is="iconComponent(item)" v-if="iconComponent(item)" :size="16" />
          <span>{{ item.label }}</span>
        </Button>
      </div>
      <Menu v-if="menuItems.length" ref="menuRef" mode="vertical" :aria-label="translate('sheet.context-menu.aria-label', '单元格上下文菜单')" :items="menuItems" :selectable="false" @click="({ key }) => execute(itemById.get(String(key)))" />
    </div>
  </Teleport>
</template>

<style scoped>
.sheet-context-menu { position: fixed; z-index: 3000; box-sizing: border-box; width: 208px; max-width: calc(100vw - 8px); max-height: min(520px, calc(100vh - 8px)); overflow: auto; border: 1px solid #d8dde5; border-radius: 6px; background: #fff; box-shadow: 0 4px 12px #17203326; color: #263244; padding: 4px 8px; }
.sheet-context-menu__quick { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); align-items: center; gap: 6px; padding: 4px 0; }
.sheet-context-menu :deep(.sheet-context-menu__quick-item) { display: flex; min-width: 0; width: 100%; height: 48px; flex-direction: column; align-items: center; justify-content: center; gap: 2px; border: 0; border-radius: 6px; padding: 4px; color: inherit; line-height: normal; }
.sheet-context-menu :deep(.sheet-context-menu__quick-item:not(:disabled):hover), .sheet-context-menu :deep(.sheet-context-menu__quick-item:focus-visible) { outline: none; background: #f1f3f5; }
.sheet-context-menu :deep(.sheet-context-menu__quick-item:disabled) { color: #9aa3af; cursor: default; }
.sheet-context-menu :deep(.sheet-context-menu__quick-item span) { max-width: 100%; overflow: hidden; font-size: 12px; line-height: 16px; text-overflow: ellipsis; white-space: nowrap; }
.sheet-context-menu :deep(.ant-menu) { max-height: inherit; overflow: auto; border-inline-end: 0; padding: 4px 0; font-size: 14px; }
.sheet-context-menu :deep(.ant-menu-item), .sheet-context-menu :deep(.ant-menu-submenu-title) { min-height: 32px; height: auto; line-height: 22px; padding-block: 5px; font-size: 14px; }
.sheet-context-menu :deep(.ant-menu-item .ant-menu-title-content), .sheet-context-menu :deep(.ant-menu-submenu-title .ant-menu-title-content) { white-space: normal; }
.sheet-context-menu :deep(.ant-menu-item-divider) { margin: 4px 3px; }
:global(.sheet-context-menu__submenu.ant-menu-submenu-popup) { z-index: 3001; }
</style>
