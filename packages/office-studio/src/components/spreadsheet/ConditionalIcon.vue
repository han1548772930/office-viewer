<script setup lang="ts">
import { computed } from 'vue'
import { ArrowDown, ArrowDownRight, ArrowRight, ArrowUp, ArrowUpRight, Check, Circle, Diamond, Flag, Frown, Minus, Signal, Smile, Square, Star, Triangle, TriangleAlert, X } from 'lucide-vue-next'
import { conditionalIconPreset } from '../../features/conditional-formatting/conditional-icon-presets.ts'
const props = defineProps<{ iconType: string; iconId: string; size?: number }>()
const preset = computed(() => conditionalIconPreset(props.iconType, props.iconId))
const icon = computed(() => {
  const value = preset.value
  if (!value) return undefined
  if (value.shape === 'arrow') return ({ '↑': ArrowUp, '↗': ArrowUpRight, '→': ArrowRight, '↘': ArrowDownRight, '↓': ArrowDown })[value.symbol]
  return ({ circle: Circle, triangle: Triangle, diamond: Diamond, check: Check, warning: TriangleAlert, cross: value.symbol === '−' ? Minus : X, flag: Flag, rating: Signal, progress: Circle, face: value.symbol === '☹' ? Frown : Smile, box: Square, star: Star })[value.shape]
})
</script>

<template><component :is="icon" v-if="icon" :size="size ?? 18" :color="preset?.color" :aria-label="`${iconType} ${Number(iconId) + 1}`" /></template>
