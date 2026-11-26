<script setup lang="ts">
import { computed, nextTick, ref, watch } from 'vue'
import { Button, Dropdown, Menu, Tooltip, theme } from 'antdv-next'
import { Crop, ImagePlus, RotateCcw, Settings2, Trash2, Layers, AlignHorizontalJustifyCenter } from 'lucide-vue-next'
import type { MenuItemType } from 'antdv-next'
import type { DocumentImageFrame } from '../../features/document/layout/document-image-layout'
import type { DocumentImage } from '../../features/document/model/document-model'
import { sceneOffset, type DocumentViewport } from '../../features/document/rendering/document-scene'
import type { DocumentLayout } from '../../features/document/layout/document-layout-types'
import type { DocumentImageArrange } from '../../features/document/images/document-image-model'
import { DOCUMENT_IMAGE_ALIGN_TYPES, type DocumentImageAlign } from '../../features/document/images/document-image-batch'

const props = defineProps<{ image?: DocumentImage; rectangle?: DocumentImageFrame; selectedCount?: number; layout: DocumentLayout; viewport: DocumentViewport; disabled?: boolean; geometryDisabled?: boolean; geometryReason?: string; arrangeDisabled?: boolean; inactive?: boolean; contextPoint?: { x: number; y: number } }>()
const emit = defineEmits<{ edit: [mode: 'edit' | 'crop' | 'replace']; remove: []; reset: []; arrange: [value: DocumentImageArrange]; align: [value: DocumentImageAlign]; focus: []; 'close-context': [] }>()
const { token } = theme.useToken()
const count = computed(() => props.selectedCount ?? (props.image ? 1 : 0)), multiple = computed(() => count.value > 1)
const arrangeOpen = ref(false), alignOpen = ref(false), contextRef = ref<HTMLDivElement>()
const toolbarStyle = computed(() => {
  const image = props.rectangle
  if (!image) return { display: 'none' }
  const scale = props.viewport.zoom / 100, offset = sceneOffset(props.layout, props.viewport)
  const x = offset.x + image.x * scale, y = offset.y + image.y * scale, width = image.width * scale, height = image.height * scale
  if (x + width < 0 || x > props.viewport.width || y + height < 0 || y > props.viewport.height) return { display: 'none' }
  return { left: `${Math.max(8, Math.min(x, props.viewport.width - 260))}px`, top: `${Math.max(8, Math.min(y - 68, props.viewport.height - 40))}px`,
    maxWidth: `${Math.max(100, props.viewport.width - 16)}px`, color: token.value.colorText, background: token.value.colorBgElevated,
    borderColor: token.value.colorBorder, boxShadow: token.value.boxShadowSecondary }
})
const arrangeItems = computed<MenuItemType[]>(() => [
  { key: 'front', label: '置于顶层', disabled: props.disabled || props.arrangeDisabled }, { key: 'forward', label: '上移一层', disabled: props.disabled || props.arrangeDisabled },
  { key: 'back', label: '置于底层', disabled: props.disabled || props.arrangeDisabled }, { key: 'backward', label: '下移一层', disabled: props.disabled || props.arrangeDisabled },
])
const alignLabels: Array<[DocumentImageAlign, string]> = [
  [DOCUMENT_IMAGE_ALIGN_TYPES.left, '左对齐'], [DOCUMENT_IMAGE_ALIGN_TYPES.center, '水平居中'], [DOCUMENT_IMAGE_ALIGN_TYPES.right, '右对齐'],
  [DOCUMENT_IMAGE_ALIGN_TYPES.top, '顶部对齐'], [DOCUMENT_IMAGE_ALIGN_TYPES.middle, '垂直居中'], [DOCUMENT_IMAGE_ALIGN_TYPES.bottom, '底部对齐'],
  [DOCUMENT_IMAGE_ALIGN_TYPES.horizon, '水平分布'], [DOCUMENT_IMAGE_ALIGN_TYPES.vertical, '垂直分布'],
]
const alignItems = computed<MenuItemType[]>(() => alignLabels.map(([value, label]) => ({ key: `align-${value}`, label,
  disabled: props.disabled || props.geometryDisabled || count.value < 2 || (value === '7' || value === '8') && count.value < 3 })))
function returnFocus() { void nextTick(() => emit('focus')) }
function closeMenus() { arrangeOpen.value = false; alignOpen.value = false }
function arrange(key: string | number) {
  if (['front', 'forward', 'back', 'backward'].includes(String(key)) && !props.disabled && !props.arrangeDisabled) {
    closeMenus(); emit('arrange', key as DocumentImageArrange); returnFocus()
  }
}
function align(key: string | number) {
  const value = alignLabels.find(([value]) => key === `align-${value}`)?.[0]
  if (!value || props.disabled || props.geometryDisabled || count.value < 2 || (value === '7' || value === '8') && count.value < 3) return
  closeMenus(); emit('align', value); returnFocus()
}
const contextItems = computed<MenuItemType[]>(() => [
  ...(!multiple.value ? [
    { key: 'edit', label: '编辑图片', disabled: props.disabled }, { key: 'crop', label: '裁剪图片', disabled: props.disabled },
    { key: 'replace', label: '替换图片', disabled: props.disabled }, { key: 'reset', label: '重置尺寸与裁剪', disabled: props.disabled || !props.image?.intrinsicWidth || !props.image?.intrinsicHeight },
  ] : []),
  { key: 'arrange', label: '排列', disabled: props.disabled || props.arrangeDisabled, children: arrangeItems.value },
  ...(multiple.value ? [{ key: 'align', label: '对齐与分布', disabled: props.disabled || props.geometryDisabled, children: alignItems.value }] : []),
  { type: 'divider' }, { key: 'remove', label: multiple.value ? '删除选中的图片' : '删除图片', disabled: props.disabled },
])
function contextAction(key: string | number) {
  emit('close-context')
  if (props.disabled) return
  if (!multiple.value && (key === 'edit' || key === 'crop' || key === 'replace')) emit('edit', key)
  else if (key === 'remove') emit('remove')
  else if (!multiple.value && key === 'reset' && props.image?.intrinsicWidth && props.image?.intrinsicHeight) emit('reset')
  else if (String(key).startsWith('align-')) align(key)
  else arrange(key)
  returnFocus()
}
function closeContext() { emit('close-context'); returnFocus() }
watch(() => [props.inactive, props.disabled, props.selectedCount, props.rectangle], closeMenus)
watch(() => props.contextPoint, value => {
  if (value) void nextTick(() => contextRef.value?.querySelector<HTMLElement>('[role="menuitem"]:not([aria-disabled="true"])')?.focus())
})
</script>

<template>
  <div v-if="image && rectangle && count && !inactive" class="document-image-toolbar" :style="toolbarStyle" role="toolbar" aria-label="图片操作" data-document-image-toolbar @pointerdown.stop>
    <span v-if="multiple" class="document-image-count" data-document-image-selection-count>已选 {{ count }} 张</span>
    <template v-if="!multiple">
      <Tooltip title="编辑图片"><Button size="small" type="text" aria-label="编辑图片" :disabled="disabled" @click="emit('edit', 'edit')"><Settings2 :size="16" /></Button></Tooltip>
      <Tooltip title="裁剪图片"><Button size="small" type="text" aria-label="裁剪图片" :disabled="disabled" @click="emit('edit', 'crop')"><Crop :size="16" /></Button></Tooltip>
    </template>
    <Tooltip :title="multiple ? '删除选中的图片' : '删除图片'"><Button size="small" type="text" :aria-label="multiple ? '删除选中的图片' : '删除图片'" :disabled="disabled" @click="emit('remove')"><Trash2 :size="16" /></Button></Tooltip>
    <template v-if="!multiple">
      <Tooltip title="恢复原始尺寸与裁剪"><Button size="small" type="text" aria-label="重置图片" :disabled="disabled || !image.intrinsicWidth || !image.intrinsicHeight" @click="emit('reset')"><RotateCcw :size="16" /></Button></Tooltip>
      <Tooltip title="替换图片"><Button size="small" type="text" aria-label="替换图片" :disabled="disabled" @click="emit('edit', 'replace')"><ImagePlus :size="16" /></Button></Tooltip>
    </template>
    <Dropdown :open="arrangeOpen" :trigger="['click']" :disabled="disabled || arrangeDisabled" :auto-focus="true" @open-change="value => { arrangeOpen = value; if (!value) returnFocus() }">
      <Tooltip :title="arrangeDisabled ? '排列需要同一内容容器中的图片' : '排列图片'"><Button size="small" type="text" aria-label="排列图片" :disabled="disabled || arrangeDisabled"><Layers :size="16" /></Button></Tooltip>
      <template #overlay><Menu :items="arrangeItems" :selectable="false" @click="value => arrange(value.key)" @keydown.esc.stop="closeMenus(); returnFocus()" /></template>
    </Dropdown>
    <Dropdown v-if="multiple" :open="alignOpen" :trigger="['click']" :disabled="disabled || geometryDisabled" :auto-focus="true" @open-change="value => { alignOpen = value; if (!value) returnFocus() }">
      <Tooltip :title="geometryDisabled ? geometryReason : '对齐与分布'"><Button size="small" type="text" aria-label="对齐与分布" :disabled="disabled || geometryDisabled"><AlignHorizontalJustifyCenter :size="16" /></Button></Tooltip>
      <template #overlay><Menu :items="alignItems" :selectable="false" @click="value => align(value.key)" @keydown.esc.stop="closeMenus(); returnFocus()" /></template>
    </Dropdown>
  </div>
  <div v-if="contextPoint && image && count && !inactive" ref="contextRef" class="document-image-context" :style="{ left: `${contextPoint.x}px`, top: `${contextPoint.y}px`, background: token.colorBgElevated, borderColor: token.colorBorder, boxShadow: token.boxShadowSecondary }"
    aria-label="图片右键菜单" @pointerdown.stop @contextmenu.prevent @keydown.esc.stop="closeContext">
    <Menu :items="contextItems" :selectable="false" @click="value => contextAction(value.key)" />
  </div>
</template>

<style scoped>
.document-image-toolbar { position: absolute; z-index: 6; display: flex; align-items: center; gap: 2px; padding: 3px; border: 1px solid; border-radius: 6px; overflow-x: auto; }
.document-image-context { position: absolute; z-index: 7; width: 200px; border: 1px solid; border-radius: 6px; }
.document-image-count { padding: 0 6px; white-space: nowrap; }
</style>
