<script setup lang="ts">
import { computed, onBeforeUnmount, onMounted, ref } from 'vue'
import { Button, Popover } from 'antdv-next'
import { ChevronDown } from 'lucide-vue-next'
import Control, { type HeaderFooterRibbonKind, type HeaderFooterRibbonState } from './DocumentHeaderFooterRibbonControl.vue'
import { headerFooterControlDisabled } from './document-header-footer-capabilities'

const props = defineProps<{ kind: HeaderFooterRibbonKind; state: HeaderFooterRibbonState; title: string; grid?: boolean; overflow?: boolean; measurement?: boolean }>()
const emit = defineEmits<{ insert: [value: string] }>()
const container = ref<HTMLElement>(), width = ref(Infinity), open = ref(false)
let observer: ResizeObserver | undefined
onMounted(() => {
  const toolbar = container.value?.closest('[role="toolbar"]') ?? container.value?.closest('.workbench-ribbon')
  if (!toolbar) return
  const measure = () => { width.value = toolbar.getBoundingClientRect().width }
  observer = new ResizeObserver(measure); observer.observe(toolbar); measure()
})
onBeforeUnmount(() => observer?.disconnect())
const compact = computed(() => !props.overflow && props.kind !== 'insertion' && width.value < { position: 1100, options: 900, navigation: 650 }[props.kind])
function forward(value: string) { open.value = false; if (!props.measurement) emit('insert', value) }
</script>
<template>
  <div ref="container" :inert="measurement || undefined">
    <Popover v-if="compact" v-model:open="open" trigger="click" placement="bottomLeft">
      <Button size="small" type="text" :disabled="headerFooterControlDisabled(kind, state)" :aria-expanded="open" @mousedown.prevent>{{ title }} <ChevronDown :size="12" /></Button>
      <template #content><Control :kind="kind" :state="state" @insert="forward" /></template>
    </Popover>
    <Control v-else :kind="kind" :state="state" :classic="!grid && !overflow" @insert="forward" />
  </div>
</template>
