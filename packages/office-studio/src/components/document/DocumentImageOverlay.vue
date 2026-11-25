<script setup lang="ts">
import { computed } from 'vue'
import { Button, Dropdown, Menu, Tooltip, theme } from 'antdv-next'
import { Crop, ImagePlus, RotateCcw, Settings2, Trash2, Layers } from 'lucide-vue-next'
import type { MenuItemType } from 'antdv-next'
import type { LayoutImage } from '../../features/document/layout/document-layout-types'
import type { DocumentImage } from '../../features/document/model/document-model'
import { sceneOffset, type DocumentViewport } from '../../features/document/rendering/document-scene'
import type { DocumentLayout } from '../../features/document/layout/document-layout-types'
import type { DocumentImageArrange } from '../../features/document/images/document-image-model'

const props = defineProps<{ image?: DocumentImage; rectangle?: LayoutImage; layout: DocumentLayout; viewport: DocumentViewport; disabled?: boolean; inactive?: boolean; contextPoint?: { x: number; y: number } }>()
const emit = defineEmits<{ edit: [mode: 'edit' | 'crop' | 'replace']; remove: []; reset: []; arrange: [value: DocumentImageArrange]; 'close-context': [] }>()
const { token } = theme.useToken()
const toolbarStyle = computed(() => {
  const image = props.rectangle
  if (!image) return { display: 'none' }
  const scale = props.viewport.zoom / 100, offset = sceneOffset(props.layout, props.viewport)
  const x = offset.x + image.x * scale, y = offset.y + image.y * scale, width = image.width * scale, height = image.height * scale
  if (x + width < 0 || x > props.viewport.width || y + height < 0 || y > props.viewport.height) return { display: 'none' }
  return { left: `${Math.max(8, Math.min(x, props.viewport.width - 260))}px`, top: `${Math.max(8, Math.min(y - 36, props.viewport.height - 40))}px`,
    maxWidth: `${Math.max(100, props.viewport.width - 16)}px`, color: token.value.colorText, background: token.value.colorBgElevated,
    borderColor: token.value.colorBorder, boxShadow: token.value.boxShadowSecondary }
})
const arrangeItems = computed<MenuItemType[]>(() => [
  { key: 'front', label: '置于顶层', disabled: props.disabled }, { key: 'forward', label: '上移一层', disabled: props.disabled },
  { key: 'back', label: '置于底层', disabled: props.disabled }, { key: 'backward', label: '下移一层', disabled: props.disabled },
])
function arrange(key: string | number) { if (['front', 'forward', 'back', 'backward'].includes(String(key)) && !props.disabled) emit('arrange', key as DocumentImageArrange) }
const contextItems = computed<MenuItemType[]>(() => [
  { key: 'edit', label: '编辑图片', disabled: props.disabled }, { key: 'crop', label: '裁剪图片', disabled: props.disabled },
  { key: 'replace', label: '替换图片', disabled: props.disabled }, { key: 'reset', label: '重置尺寸与裁剪', disabled: props.disabled || !props.image?.intrinsicWidth || !props.image?.intrinsicHeight },
  { key: 'arrange', label: '排列', children: arrangeItems.value }, { type: 'divider' }, { key: 'remove', label: '删除图片', disabled: props.disabled },
])
function contextAction(key: string | number) {
  emit('close-context')
  if (props.disabled) return
  if (key === 'edit' || key === 'crop' || key === 'replace') emit('edit', key)
  else if (key === 'remove') emit('remove')
  else if (key === 'reset' && props.image?.intrinsicWidth && props.image?.intrinsicHeight) emit('reset')
  else arrange(key)
}
</script>

<template>
  <div v-if="image && rectangle && !inactive" class="document-image-toolbar" :style="toolbarStyle" role="toolbar" aria-label="图片操作" data-document-image-toolbar @pointerdown.stop>
    <Tooltip title="编辑图片"><Button size="small" type="text" aria-label="编辑图片" :disabled="disabled" @click="emit('edit', 'edit')"><Settings2 :size="16" /></Button></Tooltip>
    <Tooltip title="裁剪图片"><Button size="small" type="text" aria-label="裁剪图片" :disabled="disabled" @click="emit('edit', 'crop')"><Crop :size="16" /></Button></Tooltip>
    <Tooltip title="删除图片"><Button size="small" type="text" aria-label="删除图片" :disabled="disabled" @click="emit('remove')"><Trash2 :size="16" /></Button></Tooltip>
    <Tooltip title="恢复原始尺寸与裁剪"><Button size="small" type="text" aria-label="重置图片" :disabled="disabled || !image.intrinsicWidth || !image.intrinsicHeight" @click="emit('reset')"><RotateCcw :size="16" /></Button></Tooltip>
    <Tooltip title="替换图片"><Button size="small" type="text" aria-label="替换图片" :disabled="disabled" @click="emit('edit', 'replace')"><ImagePlus :size="16" /></Button></Tooltip>
    <Dropdown :trigger="['click']" :disabled="disabled"><Tooltip title="排列图片"><Button size="small" type="text" aria-label="排列图片" :disabled="disabled"><Layers :size="16" /></Button></Tooltip>
      <template #overlay><Menu :items="arrangeItems" :selectable="false" @click="value => arrange(value.key)" /></template>
    </Dropdown>
  </div>
  <div v-if="contextPoint && image && !inactive" class="document-image-context" :style="{ left: `${contextPoint.x}px`, top: `${contextPoint.y}px`, background: token.colorBgElevated, borderColor: token.colorBorder, boxShadow: token.boxShadowSecondary }"
    aria-label="图片右键菜单" @pointerdown.stop @contextmenu.prevent @keydown.esc="emit('close-context')">
    <Menu :items="contextItems" :selectable="false" @click="value => contextAction(value.key)" />
  </div>
</template>

<style scoped>
.document-image-toolbar { position: absolute; z-index: 6; display: flex; align-items: center; gap: 2px; padding: 3px; border: 1px solid; border-radius: 6px; overflow-x: auto; }
.document-image-context { position: absolute; z-index: 7; width: 200px; border: 1px solid; border-radius: 6px; }
</style>
