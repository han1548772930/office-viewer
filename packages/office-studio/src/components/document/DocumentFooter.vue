<script setup lang="ts">
import { ChevronLeft, ChevronRight, Maximize, Minus, Plus } from 'lucide-vue-next'
import { Button, Select, Slider, Tooltip, theme } from 'antdv-next'
import { computed } from 'vue'
import { DOCUMENT_ZOOM_SHORTCUTS } from '../../features/document/services/document-zoom'

defineProps<{ activePage: number; pageCount: number; characterCount: number; selectedCharacters: number; zoom: number; zoomDisabled?: boolean; modified: boolean; focusMode: boolean }>()
const emit = defineEmits<{ navigate: [direction: -1 | 1]; zoom: [value: number]; 'exit-focus': [] }>()
const zoomOptions = DOCUMENT_ZOOM_SHORTCUTS.map(value => ({ label: `${value}%`, value }))
const { token } = theme.useToken()
const footerStyle = computed(() => ({ '--document-footer-border': token.value.colorBorderSecondary,
  '--document-footer-text': token.value.colorTextSecondary, '--document-footer-background': token.value.colorBgContainer,
  '--document-footer-disabled': token.value.colorTextDisabled, '--document-footer-warning': token.value.colorWarningText }))
</script>

<template>
  <footer class="document-statusbar" :style="footerStyle" aria-label="文档状态栏">
    <div class="status-left">
      <span class="page-status">第 {{ Math.max(activePage + 1, 1) }} 页，共 {{ Math.max(pageCount, 1) }} 页</span>
      <span class="status-divider" aria-hidden="true"></span>
      <span class="character-status">{{ selectedCharacters ? `已选 ${selectedCharacters} / ` : '' }}{{ characterCount }} 个字符</span>
      <span class="save-status" :class="{ modified }" role="status">{{ modified ? '未保存更改' : '无未保存更改' }}</span>
    </div>
    <div class="status-right">
      <Tooltip v-if="focusMode" title="退出专注模式"><Button type="text" aria-label="退出专注模式" @click="emit('exit-focus')"><Maximize :size="15" /></Button></Tooltip>
      <Tooltip title="上一页"><Button type="text" aria-label="上一页" :disabled="activePage <= 0" @click="emit('navigate', -1)"><ChevronLeft :size="15" /></Button></Tooltip>
      <Tooltip title="下一页"><Button type="text" aria-label="下一页" :disabled="activePage >= pageCount - 1" @click="emit('navigate', 1)"><ChevronRight :size="15" /></Button></Tooltip>
      <span class="status-divider" aria-hidden="true"></span>
      <Tooltip title="缩小"><Button type="text" aria-label="缩小" :disabled="zoomDisabled || zoom <= 10" @click="emit('zoom', zoom - 10)"><Minus :size="14" /></Button></Tooltip>
      <Slider class="status-zoom" :min="10" :max="400" :step="1" :value="zoom" :disabled="zoomDisabled" :tooltip="{ formatter: (value?: number) => `${value}%` }" aria-label="缩放比例" @change="value => { if (typeof value === 'number') emit('zoom', value) }" />
      <Tooltip title="放大"><Button type="text" aria-label="放大" :disabled="zoomDisabled || zoom >= 400" @click="emit('zoom', zoom + 10)"><Plus :size="14" /></Button></Tooltip>
      <Select class="zoom-value" :value="zoom" :options="zoomOptions" :disabled="zoomDisabled" size="small" aria-label="文档缩放" @change="emit('zoom', Number($event))" />
    </div>
  </footer>
</template>

<style scoped>
.document-statusbar { box-sizing: border-box; flex: 0 0 36px; min-width: 0; height: 36px; display: flex; align-items: center; justify-content: space-between; gap: 10px; padding: 0 12px; border-top: 1px solid var(--document-footer-border); color: var(--document-footer-text); background: var(--document-footer-background); font-size: 12px; }
.status-left, .status-right { display: flex; align-items: center; gap: 8px; min-width: 0; }
.status-left { overflow: hidden; }
.page-status, .character-status, .save-status { white-space: nowrap; }
.status-divider { flex: 0 0 1px; width: 1px; height: 14px; background: var(--document-footer-border); }
.status-right { flex-shrink: 0; gap: 2px; }
.status-right :deep(.ant-btn) { display: grid; place-items: center; width: 24px; height: 24px; padding: 0; color: var(--document-footer-text); }
.status-right :deep(.ant-btn:disabled) { color: var(--document-footer-disabled); }
.status-right :deep(.status-zoom) { width: 90px; margin: 0 5px; }
.zoom-value { width: 76px; }
.save-status { margin-left: 4px; color: var(--document-footer-text); }
.save-status.modified { color: var(--document-footer-warning); }
@media (max-width: 900px) { .save-status { display: none; } }
@media (max-width: 650px) { .document-statusbar { gap: 6px; padding: 0 6px; font-size: 11px; } .status-left { gap: 5px; } .status-zoom { display: none; } .status-right { gap: 0; } .character-status { overflow: hidden; text-overflow: ellipsis; } }
</style>
