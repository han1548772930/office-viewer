<script setup lang="ts">
import { computed } from 'vue'
import { Button } from 'antdv-next'
import { listKind, listMarker, listPresets, type ListKind, type ListPreset } from '../../features/document/model/document-lists'

const props = defineProps<{ kind: ListKind; active?: ListPreset; title: string }>()
const emit = defineEmits<{ insert: [value: ListPreset] }>()
const cards = computed(() => listPresets.filter(p => listKind(p) === props.kind).map(preset => ({
  preset, preview: [0, 1, 2, 0].map(level => ({ level, glyph: listMarker(preset, level) })),
})))
</script>

<template>
  <div class="doc-list-picker" role="group" :aria-label="title">
    <Button v-for="card in cards" :key="card.preset" class="doc-list-card" :aria-pressed="active === card.preset"
      :aria-label="`${title}: ${card.preview.slice(0, 3).map(row => row.glyph).join(' ')}`" :data-list-preset="card.preset" @mousedown.prevent @click="emit('insert', card.preset)">
      <span class="doc-list-preview" aria-hidden="true">
        <span v-for="(row, index) in card.preview" :key="index" class="doc-list-row" :style="{ paddingInlineStart: `${row.level * 8}px` }">
          <span>{{ row.glyph }}</span><i />
        </span>
      </span>
    </Button>
  </div>
</template>

<style scoped>
.doc-list-picker { display: grid; grid-template-columns: repeat(3, minmax(0, 72px)); gap: 8px; padding: 6px; }
.doc-list-card { width: 100%; height: 80px; padding: 0; overflow: hidden; border-radius: 4px; background: var(--ribbon-surface); color: var(--ribbon-text); }
.doc-list-card[aria-pressed="true"], .doc-list-card:hover { border-color: var(--ribbon-primary); }
.doc-list-preview { box-sizing: border-box; width: 100%; height: 100%; padding: 8px; display: grid; grid-template-rows: repeat(4, 1fr); gap: 4px; }
.doc-list-row { display: flex; align-items: center; gap: 4px; min-width: 0; font-size: 10px; line-height: 1; }
.doc-list-row > span { min-width: 12px; text-align: center; }
.doc-list-row i { min-width: 2px; flex: 1; height: 4px; border-radius: 2px; background: var(--ribbon-border); }
</style>
