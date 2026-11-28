<script setup lang="ts">
import { computed, inject, shallowRef, watch } from 'vue'
import { Button, Tooltip } from 'antdv-next'
import { ChevronDown, ChevronRight } from 'lucide-vue-next'
import type { DocumentModel } from '../../domain/model'
import type { DocumentCommand } from '../../features/document/commands/document-commands'
import { documentHeadingCollapseState, documentHeadingVisibility, SET_DOCUMENT_HEADING_COLLAPSED_COMMAND, type DocumentHeadingSection } from '../../features/document/headings/heading-collapse'
import type { DocumentInputSession } from '../../features/document/model/document-input-session'
import { runText } from '../../features/document/model/document-model'
import type { DocumentLayout } from '../../features/document/layout/document-layout'
import type { DocumentPermissionContext } from '../../features/document/permissions/document-permissions'
import { sceneOffset, type DocumentViewport } from '../../features/document/rendering/document-scene'
import { OfficeTranslatorKey } from '../../i18n/office-i18n'

const props = defineProps<{ model: DocumentModel; layout: DocumentLayout; viewport: DocumentViewport; disabled?: boolean;
  permission?: DocumentPermissionContext; captureInputSession?: (source?: DocumentModel) => DocumentInputSession | undefined }>()
const emit = defineEmits<{ command: [command: DocumentCommand] }>()
const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const visibility = computed(() => documentHeadingVisibility(props.model))
const context = computed(() => props.permission ?? { documentId: props.model.id ?? '', userId: 'local-user' })
const scale = computed(() => props.viewport.zoom / 100)
const tooltipId = shallowRef<string>()
const pressed = shallowRef<{ source: DocumentModel; id: string; collapsed: boolean; generation?: number }>()
const entries = computed(() => {
  const firstLines = new Map(props.layout.lines.filter(line => !line.host && !line.repeatedHeader && line.start === 0)
    .map(line => [line.target.cellBlockId ?? line.target.blockId, line]))
  const offset = sceneOffset(props.layout, props.viewport), k = scale.value
  return visibility.value.sections.flatMap(section => {
    if (section.endIndex <= section.index + 1 || visibility.value.hiddenOwners.has(section.paragraphId)) return []
    const line = firstLines.get(section.paragraphId)
    if (!line) return []
    const left = offset.x + (line.x - 24) * k, top = offset.y + (line.y + line.height / 2 - 14) * k
    if (left + 24 * k <= 0 || left >= props.viewport.width || top + 28 * k <= 0 || top >= props.viewport.height) return []
    const state = documentHeadingCollapseState(props.model, section.paragraphId, !section.collapsed, context.value, visibility.value)
    const label = t(section.collapsed ? 'document.heading.expand' : 'document.heading.collapse', section.collapsed ? '展开' : '收起')
    return [{ section, state, label, ariaLabel: `${label} ${runText(section.block.runs)}`,
      style: { left: `${Math.max(0, left)}px`, top: `${top}px`, width: `${24 * k}px`, height: `${28 * k}px` } }]
  })
})
function begin(event: PointerEvent, section: DocumentHeadingSection) {
  pressed.value = undefined
  if (event.button !== 0 || props.disabled) return
  const input = props.captureInputSession?.(props.model)
  if (props.captureInputSession && !input?.valid) return
  pressed.value = { source: props.model, id: section.paragraphId, collapsed: section.collapsed, generation: input?.generation }
}
function activate(event: MouseEvent, section: DocumentHeadingSection) {
  const owner = pressed.value
  pressed.value = undefined
  if (props.disabled || event.detail > 0 && (!owner || owner.source !== props.model || owner.id !== section.paragraphId || owner.collapsed !== section.collapsed)) return
  const input = props.captureInputSession?.(props.model)
  if (props.captureInputSession && (!input?.valid || owner && input.generation !== owner.generation)) return
  if (!documentHeadingCollapseState(props.model, section.paragraphId, !section.collapsed, context.value, visibility.value).available) return
  tooltipId.value = undefined
  emit('command', { id: SET_DOCUMENT_HEADING_COLLAPSED_COMMAND, paragraphId: section.paragraphId, collapsed: !section.collapsed })
}
watch([() => props.model, () => props.disabled, () => props.viewport.scrollLeft, () => props.viewport.scrollTop, () => props.viewport.zoom],
  () => { pressed.value = undefined; tooltipId.value = undefined })
</script>

<template>
  <div class="document-heading-collapse-overlay">
    <Tooltip v-for="entry in entries" :key="entry.section.paragraphId" :title="entry.state.reason ?? entry.label"
      :open="tooltipId === entry.section.paragraphId" :mouse-enter-delay="0.5" :trigger="['hover', 'focus']"
      @open-change="open => tooltipId = open ? entry.section.paragraphId : undefined">
      <Button class="document-heading-collapse-button" type="text" size="small" :style="entry.style"
        :data-heading-collapse="entry.section.paragraphId" :aria-label="entry.ariaLabel" :aria-expanded="!entry.section.collapsed"
        :disabled="disabled || !entry.state.available" @pointerdown.prevent.stop="begin($event, entry.section)"
        @pointercancel="pressed = undefined" @pointerleave="pressed = undefined" @click.prevent.stop="activate($event, entry.section)" @dblclick.stop>
        <ChevronRight v-if="entry.section.collapsed" :size="14 * scale" aria-hidden="true" />
        <ChevronDown v-else :size="14 * scale" aria-hidden="true" />
      </Button>
    </Tooltip>
  </div>
</template>

<style scoped>
.document-heading-collapse-overlay { position: absolute; inset: 0; pointer-events: none; z-index: 4; }
.document-heading-collapse-button { position: absolute; display: flex; align-items: center; justify-content: center; padding: 0; pointer-events: auto; }
</style>
