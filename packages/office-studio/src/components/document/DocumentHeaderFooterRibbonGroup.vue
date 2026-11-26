<script setup lang="ts">
import { computed, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import { Button, Popover } from 'antdv-next'
import { ChevronDown } from 'lucide-vue-next'
import Control, { type HeaderFooterRibbonKind, type HeaderFooterRibbonState } from './DocumentHeaderFooterRibbonControl.vue'
import { headerFooterControlDisabled } from './document-header-footer-capabilities'
import { useRibbonTheme } from '../../platform/ribbon/ribbon-theme'

const props = defineProps<{ kind: HeaderFooterRibbonKind; state: HeaderFooterRibbonState; title: string; grid?: boolean; overflow?: boolean; measurement?: boolean }>()
const emit = defineEmits<{ insert: [value: string] }>()
const themeStyle = useRibbonTheme()
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
const controlDisabled = computed(() => headerFooterControlDisabled(props.kind, props.state))
watch([() => props.state.requestScope, () => props.state.visible, controlDisabled, compact], ([scope, visible, disabled, isCompact], [previousScope]) => {
  if (scope !== previousScope || !visible || disabled || !isCompact) open.value = false
}, { flush: 'sync' })
function forward(value: string) { open.value = false; if (!props.measurement) emit('insert', value) }
</script>
<template>
  <div ref="container" class="header-footer-ribbon-group" :class="{ 'is-grid': grid && !overflow, 'is-overflow': overflow }" :style="themeStyle" :inert="measurement || undefined">
    <Popover v-if="compact" v-model:open="open" trigger="click" placement="bottomLeft">
      <Button class="header-footer-ribbon-group__trigger" size="small" type="text" :styles="{ root: { border: 0 } }" :disabled="controlDisabled" :aria-expanded="open" @mousedown.prevent><span class="header-footer-ribbon-group__caption"><span>{{ title }}</span><ChevronDown :size="12" /></span></Button>
      <template #content><Control :kind="kind" :state="state" @insert="forward" /></template>
    </Popover>
    <Control v-else :kind="kind" :state="state" :classic="!grid && !overflow" @insert="forward" />
  </div>
</template>

<style scoped>
.header-footer-ribbon-group { display: flex; width: max-content; min-width: 0; height: var(--ribbon-selector-height, 28px); align-items: center; }
.header-footer-ribbon-group.is-grid { width: 100%; height: 100%; justify-content: center; }
.header-footer-ribbon-group.is-overflow { width: 100%; height: auto; }
.header-footer-ribbon-group__trigger { box-sizing: border-box; height: var(--ribbon-selector-height, 28px); padding: 0 6px; font-size: 14px; line-height: 20px; color: var(--ribbon-text); border-radius: 4px; }
.is-grid .header-footer-ribbon-group__trigger { width: 100%; height: var(--ribbon-grid-row-height, 24px); font-size: 12px; line-height: 16px; }
.header-footer-ribbon-group__caption { display: inline-flex; min-width: 0; max-width: 100%; align-items: center; justify-content: center; gap: 4px; }
.header-footer-ribbon-group__caption > span { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.header-footer-ribbon-group__caption :deep(svg) { display: block; width: 12px; height: 12px; flex: 0 0 12px; }
.header-footer-ribbon-group__trigger:not(:disabled):hover { background: var(--ribbon-hover-background); }
.header-footer-ribbon-group__trigger:disabled { color: var(--ribbon-disabled); background: transparent; }
.header-footer-ribbon-group__trigger:focus-visible { outline: 2px solid var(--ribbon-primary); outline-offset: -2px; }
</style>
