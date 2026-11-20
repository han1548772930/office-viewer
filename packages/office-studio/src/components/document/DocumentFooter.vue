<script setup lang="ts">
import { ChevronLeft, ChevronRight, Maximize, Minus, Plus } from 'lucide-vue-next'
import { Button, Select, Tooltip } from 'antdv-next'

defineProps<{ activePage: number; pageCount: number; characterCount: number; selectedCharacters: number; zoom: number; modified: boolean; focusMode: boolean }>()
const emit = defineEmits<{ navigate: [direction: -1 | 1]; zoom: [value: number]; 'exit-focus': [] }>()
const zoomOptions = [25, 50, 75, 100, 125, 150, 200, 300, 400].map((value) => ({ label: `${value}%`, value }))
</script>

<template>
  <footer class="document-statusbar" aria-label="文档状态栏">
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
      <Tooltip title="缩小"><Button type="text" aria-label="缩小" :disabled="zoom <= 10" @click="emit('zoom', zoom - 10)"><Minus :size="14" /></Button></Tooltip>
      <input class="status-zoom" type="range" min="10" max="400" :value="zoom" aria-label="缩放比例" @input="emit('zoom', Number(($event.target as HTMLInputElement).value))">
      <Tooltip title="放大"><Button type="text" aria-label="放大" :disabled="zoom >= 400" @click="emit('zoom', zoom + 10)"><Plus :size="14" /></Button></Tooltip>
      <Select class="zoom-value" :value="zoom" :options="zoomOptions" size="small" aria-label="文档缩放" @change="emit('zoom', Number($event))" />
    </div>
  </footer>
</template>

<style scoped>
.document-statusbar { box-sizing: border-box; flex: 0 0 36px; min-width: 0; height: 36px; display: flex; align-items: center; justify-content: space-between; gap: 10px; padding: 0 12px; border-top: 1px solid #e4e7eb; color: #69737f; background: #fff; font-size: 12px; }
.status-left, .status-right { display: flex; align-items: center; gap: 8px; min-width: 0; }
.status-left { overflow: hidden; }
.page-status, .character-status, .save-status { white-space: nowrap; }
.status-divider { flex: 0 0 1px; width: 1px; height: 14px; background: #e4e7eb; }
.status-right { flex-shrink: 0; gap: 2px; }
.status-right :deep(.ant-btn) { display: grid; place-items: center; width: 24px; height: 24px; padding: 0; color: #606b77; }
.status-right :deep(.ant-btn:disabled) { color: #b7bec8; }
.status-zoom { width: 90px; height: 20px; margin: 0 5px; accent-color: #367ddd; }
.zoom-value { width: 76px; }
.save-status { margin-left: 4px; color: #9299a3; }
.save-status.modified { color: #a46a18; }
@media (max-width: 900px) { .save-status { display: none; } }
@media (max-width: 650px) { .document-statusbar { gap: 6px; padding: 0 6px; font-size: 11px; } .status-left { gap: 5px; } .status-zoom { display: none; } .status-right { gap: 0; } .character-status { overflow: hidden; text-overflow: ellipsis; } }
</style>
