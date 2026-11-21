<script setup lang="ts">
import { computed } from 'vue'
import { Button, Input, Modal, Switch } from 'antdv-next'
import type { DocumentModel } from '../../domain/model'
import { documentSectionIds, resolveSectionHeaderFooter, type SectionHeaderFooterPatch } from '../../features/document/services/section-header-footer'

const props = defineProps<{ open: boolean; model: DocumentModel; sectionId: string }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; apply: [patch: SectionHeaderFooterPatch] }>()
const section = computed(() => resolveSectionHeaderFooter(props.model, props.sectionId))
const sections = computed(() => documentSectionIds(props.model))

function update(key: keyof typeof section.value, value: string | boolean) {
  emit('apply', { [key]: value })
}
</script>

<template>
  <Modal :open="open" title="页眉和页脚" ok-text="应用" cancel-text="关闭" :width="520" @ok="emit('update:open', false)" @cancel="emit('update:open', false)">
    <div class="header-footer-panel">
      <label class="header-footer-panel__row"><span>当前节</span><Input :value="`${sections.indexOf(sectionId) + 1}`" readonly /></label>
      <label class="header-footer-panel__toggle"><span>链接到前一节</span><Switch :checked="section.linkToPrevious" :disabled="sections.indexOf(sectionId) === 0" @change="update('linkToPrevious', $event)" /></label>
      <label class="header-footer-panel__row"><span>页眉</span><textarea :value="section.header" :disabled="section.linkToPrevious" maxlength="500" rows="1" @input="update('header', ($event.target as HTMLTextAreaElement).value)" /></label>
      <label class="header-footer-panel__row"><span>页脚</span><textarea :value="section.footer" :disabled="section.linkToPrevious" maxlength="500" rows="1" @input="update('footer', ($event.target as HTMLTextAreaElement).value)" /></label>
      <label class="header-footer-panel__toggle"><span>首页不同</span><Switch :checked="section.differentFirstPage" :disabled="section.linkToPrevious" @change="update('differentFirstPage', $event)" /></label>
      <template v-if="section.differentFirstPage">
        <label class="header-footer-panel__row"><span>首页页眉</span><textarea :value="section.firstPageHeader" :disabled="section.linkToPrevious" maxlength="500" rows="1" @input="update('firstPageHeader', ($event.target as HTMLTextAreaElement).value)" /></label>
        <label class="header-footer-panel__row"><span>首页页脚</span><textarea :value="section.firstPageFooter" :disabled="section.linkToPrevious" maxlength="500" rows="1" @input="update('firstPageFooter', ($event.target as HTMLTextAreaElement).value)" /></label>
      </template>
      <label class="header-footer-panel__toggle"><span>奇偶页不同</span><Switch :checked="section.differentOddEven" :disabled="section.linkToPrevious" @change="update('differentOddEven', $event)" /></label>
      <template v-if="section.differentOddEven">
        <label class="header-footer-panel__row"><span>偶数页页眉</span><textarea :value="section.evenPageHeader" :disabled="section.linkToPrevious" maxlength="500" rows="1" @input="update('evenPageHeader', ($event.target as HTMLTextAreaElement).value)" /></label>
        <label class="header-footer-panel__row"><span>偶数页页脚</span><textarea :value="section.evenPageFooter" :disabled="section.linkToPrevious" maxlength="500" rows="1" @input="update('evenPageFooter', ($event.target as HTMLTextAreaElement).value)" /></label>
      </template>
      <Button type="text" @click="emit('apply', { header: '', footer: '', firstPageHeader: '', firstPageFooter: '', evenPageHeader: '', evenPageFooter: '' })">清除本节页眉页脚内容</Button>
    </div>
  </Modal>
</template>

<style scoped>
.header-footer-panel { display: grid; gap: 12px; }
.header-footer-panel__row, .header-footer-panel__toggle { display: grid; grid-template-columns: 140px minmax(0, 1fr); align-items: center; gap: 12px; color: #465365; font-size: 13px; }
.header-footer-panel__toggle :deep(.ant-switch) { justify-self: start; }
.header-footer-panel textarea { box-sizing: border-box; width: 100%; min-width: 0; min-height: 32px; resize: vertical; padding: 5px 8px; border: 1px solid #d9dfe7; border-radius: 4px; color: #30343b; font: inherit; }
.header-footer-panel textarea:focus-visible { outline: 2px solid #8bb4f0; outline-offset: 1px; }
.header-footer-panel textarea:disabled { color: #8993a0; background: #f4f6f8; }
@media (max-width: 560px) { .header-footer-panel__row, .header-footer-panel__toggle { grid-template-columns: minmax(0, 1fr); gap: 5px; } }
</style>
