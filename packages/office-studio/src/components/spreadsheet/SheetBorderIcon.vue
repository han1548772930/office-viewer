<script setup lang="ts">
import type { SheetBorderPreset } from '../../features/sheet-workbench/sheet-toolbar-format.ts'

withDefaults(defineProps<{ preset: SheetBorderPreset; size?: number }>(), { size: 18 })
const diagonalPaths: Partial<Record<SheetBorderPreset, string>> = {
  'diagonal-down': 'M0 0 L20 20',
  'diagonal-down-pair': 'M20 10 L0 0 L10 20',
  'diagonal-down-triple': 'M20 10 L0 0 L10 20 M0 0 L20 20',
  'diagonal-up': 'M0 20 L20 0',
  'diagonal-up-pair': 'M0 10 L20 0 L10 20',
}
</script>

<template>
  <span class="border-preview" :class="`is-${preset}`" :data-border-preset="preset" :style="{ width: `${size}px`, height: `${size}px` }" aria-hidden="true">
    <svg v-if="diagonalPaths[preset]" viewBox="0 0 20 20"><path :d="diagonalPaths[preset]" fill="none" stroke="currentColor" /></svg>
  </span>
</template>

<style scoped>
.border-preview { box-sizing: border-box; display: block; flex: 0 0 auto; border: 1px solid var(--edge-guide, var(--ribbon-border)); color: var(--edge-ink, currentColor); }
.border-preview svg { display: block; width: 100%; height: 100%; }
.border-preview.is-all { border-color: currentColor; }
.border-preview.is-outside { border: 2px solid currentColor; }
.border-preview.is-all, .border-preview.is-inside { background: linear-gradient(90deg, transparent 46%, currentColor 47%, currentColor 53%, transparent 54%), linear-gradient(transparent 46%, currentColor 47%, currentColor 53%, transparent 54%); }
.border-preview.is-top { border-top: 2px solid currentColor; }
.border-preview.is-bottom { border-bottom: 2px solid currentColor; }
.border-preview.is-left { border-left: 2px solid currentColor; }
.border-preview.is-right { border-right: 2px solid currentColor; }
.border-preview.is-horizontal { background: linear-gradient(transparent 46%, currentColor 47%, currentColor 53%, transparent 54%); }
.border-preview.is-vertical { background: linear-gradient(90deg, transparent 46%, currentColor 47%, currentColor 53%, transparent 54%); }
.border-preview.is-none { border-style: dashed; }
</style>
