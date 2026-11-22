<script setup lang="ts">
import { computed, ref } from 'vue'
import RibbonItemControl from './RibbonItem.vue'
import { groupRibbonGridOverflowItems, type RibbonGroup, type RibbonItem } from './ribbon-schema'
import { provideRibbonPopupScope } from './ribbon-popup-scope'
import { useRibbonTheme } from './ribbon-theme'

const props = defineProps<{ groups: RibbonGroup[]; grid?: boolean; direction: 'ltr' | 'rtl'; label: string }>()
const emit = defineEmits<{ execute: [id: string, params?: unknown]; close: [] }>()
const root = ref<HTMLElement>()
const scope = provideRibbonPopupScope(computed(() => props.direction))
const style = useRibbonTheme()
const sections = computed(() => props.groups.map(group => ({ group, rows: props.grid ? groupRibbonGridOverflowItems(group.items) : [group.items] })))
function execute(item: RibbonItem, id: string, params?: unknown) {
  if (!item.disabled && !item.hidden) emit('execute', id, params)
}
function focus() {
  const first = root.value?.querySelector<HTMLElement>('button:not(:disabled), [role="combobox"]:not(:disabled), [tabindex="0"]')
  const target = first ?? root.value
  target?.focus({ preventScroll: true })
}
function escape(event: KeyboardEvent) {
  if (event.defaultPrevented) return
  event.stopPropagation(); event.preventDefault()
  if (scope.activeId.value) scope.close()
  else emit('close')
}
defineExpose({ focus })
</script>

<template>
  <div ref="root" class="ribbon-overflow-panel" :dir="direction" :style="style" role="dialog" :aria-label="label" tabindex="-1" @keydown.esc="escape">
    <section v-for="{ group, rows } in sections" :key="group.id" :aria-label="group.title">
      <h3 v-if="group.title">{{ group.title }}</h3>
      <div v-for="(row, index) in rows" :key="index" class="ribbon-overflow-group">
        <template v-for="item in row" :key="item.id">
          <span v-if="$slots[`item-${item.id}`]" :data-ribbon-item="item.id"><slot :name="`item-${item.id}`" :item="item" :grid="grid" :overflow="true" :execute="(id: string, params?: unknown) => execute(item, id, params)" /></span>
          <RibbonItemControl v-else :item="item" overflow grid :show-label="item.gridLayout?.showLabel" @execute="(id, params) => execute(item, id, params)" />
        </template>
      </div>
    </section>
  </div>
</template>

<style scoped>
.ribbon-overflow-panel { box-sizing: border-box; inline-size: min(320px, calc(100vw - 48px)); max-block-size: min(520px, 70vh); overflow: auto; padding: 4px 8px; color: var(--ribbon-text); }
.ribbon-overflow-panel section { padding-block: 6px; }
.ribbon-overflow-panel section + section { border-block-start: 1px solid var(--ribbon-border); }
.ribbon-overflow-panel h3 { margin: 0 0 4px; padding-inline: 4px; font-size: 12px; font-weight: 500; color: var(--ribbon-text-secondary); }
.ribbon-overflow-group { display: flex; flex-wrap: wrap; align-items: center; gap: 4px; }
.ribbon-overflow-group + .ribbon-overflow-group { margin-block-start: 4px; }
.ribbon-overflow-group > * { max-inline-size: 100%; }
</style>
