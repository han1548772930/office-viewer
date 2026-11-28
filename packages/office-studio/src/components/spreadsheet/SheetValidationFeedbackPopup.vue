<script setup lang="ts">
import { computed, inject } from 'vue'
import { Alert, Button, Popover } from 'antdv-next'
import type { SheetValidationCellFeedback } from '../../features/sheet-validation/validation-cell-feedback.ts'
import type { RenderRect } from '../../rendering/contracts.ts'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = defineProps<{ feedback: SheetValidationCellFeedback; position: RenderRect; mode: 'prompt' | 'invalid'; inputOwner: string;
  composing: boolean; dismiss?: () => void }>()
const translate = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const anchorStyle = computed(() => ({ left: `${props.position.x + props.position.width}px`, top: `${props.position.y}px` }))
const promptTitle = computed(() => props.feedback.prompt?.title ?? translate('sheet.validation.prompt-title', '数据验证'))
function onKeydown(event: KeyboardEvent): void {
  if (event.key !== 'Escape' || event.isComposing || event.keyCode === 229 || props.composing || !props.dismiss) return
  event.preventDefault(); event.stopPropagation(); props.dismiss()
}
</script>

<template>
  <Popover :open="true" :trigger="[]" placement="rightTop" :auto-adjust-overflow="true" :destroy-on-hidden="true">
    <span class="sheet-validation-feedback-anchor" :style="anchorStyle" aria-hidden="true" />
    <template #content>
      <div class="sheet-validation-feedback-popup" :data-sheet-validation-feedback="`${mode}:${feedback.key}`"
        :data-ribbon-editor-owner="inputOwner" @pointerdown.stop @keydown="onKeydown">
        <template v-if="mode === 'prompt' && feedback.prompt">
          <Alert type="info" :title="promptTitle" :description="feedback.prompt.message" show-icon />
          <Button type="text" size="small" :disabled="composing" @pointerdown.prevent @click="dismiss?.()">
            {{ translate('sheet.validation.dismiss-prompt', '关闭输入提示') }}
          </Button>
        </template>
        <Alert v-for="issue in mode === 'invalid' ? feedback.issues : []" :key="issue.ruleId" type="error"
          :title="issue.title ?? translate('sheet.validation.invalid-title', '无效值')"
          :description="issue.message ?? translate(`sheet.validation.invalid.${issue.type}`, '此单元格的内容违反了验证规则。')" show-icon />
      </div>
    </template>
  </Popover>
</template>

<style scoped>
.sheet-validation-feedback-anchor { position: absolute; width: 1px; height: 1px; pointer-events: none; }
.sheet-validation-feedback-popup { display: flex; flex-direction: column; gap: 6px; width: 240px; max-width: calc(100vw - 32px); }
</style>
