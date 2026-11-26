<script setup lang="ts">
import { Button, Divider, theme } from 'antdv-next'
import { Pencil, Trash2 } from 'lucide-vue-next'

defineProps<{ presets: { id: string; values: string[] }[]; title: string; editDropdown: string; clearDropdown: string; disabled?: boolean }>()
const emit = defineEmits<{ insert: [action: string] }>()
const { token } = theme.useToken()
</script>

<template>
  <div class="quick-dropdown" :aria-label="title" role="group" :style="{ '--preset-surface': token.colorFillAlter }">
    <strong>{{ title }}</strong>
    <div class="quick-presets">
      <Button v-for="preset in presets" :key="preset.id" type="text" class="quick-preset" :styles="{ root: { border: 0 } }" :disabled="disabled" :aria-label="preset.values.join(' / ')" @mousedown.prevent @click="emit('insert', preset.id)">
        <span class="preset-values"><span v-for="value in preset.values" :key="value">{{ value }}</span></span>
      </Button>
    </div>
    <Divider class="quick-divider" />
    <Button type="text" block class="quick-action" :disabled="disabled" :classes="{ content: 'quick-action-label' }" :aria-label="editDropdown" @mousedown.prevent @click="emit('insert', 'edit-dropdown')"><Pencil :size="16" />{{ editDropdown }}</Button>
    <Button type="text" block class="quick-action" :disabled="disabled" :classes="{ content: 'quick-action-label' }" :aria-label="clearDropdown" @mousedown.prevent @click="emit('insert', 'clear-dropdown')"><Trash2 :size="16" />{{ clearDropdown }}</Button>
  </div>
</template>

<style scoped>
.quick-dropdown { width: min(360px, calc(100vw - 64px)); }
.quick-presets { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 6px; margin-block-start: 8px; }
.quick-preset { box-sizing: border-box; min-width: 0; min-height: 80px; height: auto; padding: 12px; align-items: flex-start; justify-content: flex-start; text-align: start; font-size: 14px; line-height: 20px; background: var(--preset-surface); }
.preset-values { display: grid; width: 100%; gap: 6px; white-space: normal; overflow-wrap: anywhere; }
.quick-divider { margin-block: 6px; }
.quick-action { justify-content: flex-start; }
.quick-action :deep(.quick-action-label) { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.quick-action :deep(svg) { display: block; flex: 0 0 16px; width: 16px; height: 16px; }
</style>
