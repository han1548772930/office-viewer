<script setup lang="ts">
import { computed, onDeactivated, onMounted, onUnmounted } from 'vue'
import { Dropdown } from 'antdv-next'
import type { PivotChartAxisMenuEntry } from '../services/pivot-chart-axis-menu.ts'

const props = defineProps<{ open: boolean; entries: readonly PivotChartAxisMenuEntry[]; execute(key: string): Promise<boolean> }>()
const emit = defineEmits<{ (event: 'close', restore: boolean): void }>()
const items = computed(() => [{ key: 'expand-collapse', label: '展开/折叠', children: props.entries.map(entry => ({
  key: entry.key, label: entry.label, disabled: entry.disabled,
})) }])
function keyboard(event: KeyboardEvent): void {
  if (props.open && event.key === 'Escape' && !event.isComposing && event.keyCode !== 229) {
    event.preventDefault(); event.stopPropagation(); emit('close', true)
  }
}
function choose({ key }: { key: string | number }): void { void props.execute(String(key)) }
onMounted(() => document.addEventListener('keydown', keyboard, true))
onUnmounted(() => document.removeEventListener('keydown', keyboard, true))
onDeactivated(() => emit('close', false))
</script>

<template>
  <Dropdown :open="open" :trigger="[]" :menu="{ items, selectable: false }" placement="bottomLeft" auto-focus destroy-on-hidden
    @update:open="value => { if (!value) emit('close', false) }" @menu-click="choose">
    <span class="pivot-chart-axis-menu-anchor" aria-hidden="true" />
  </Dropdown>
</template>

<style scoped>
.pivot-chart-axis-menu-anchor { display: block; width: 1px; height: 1px; pointer-events: none; }
</style>
