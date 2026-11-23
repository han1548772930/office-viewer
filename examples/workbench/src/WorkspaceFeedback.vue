<script setup lang="ts">
import { Alert, Button, Spin, theme } from 'antdv-next'
import { X } from 'lucide-vue-next'
import { useI18n } from 'vue-i18n'
import type { NoticeType } from './workspace-file-operations'

defineProps<{ parserStatus: string; parserMessage: string; notice: { type: NoticeType; message: string } | null }>()
const emit = defineEmits<{ cancelImport: []; dismissParserError: [] }>()
const { t } = useI18n()
const { token } = theme.useToken()
</script>

<template>
  <div v-if="parserStatus === 'loading'" class="workspace-task-strip" role="status" aria-live="polite" :style="{ backgroundColor: token.colorInfoBg, borderColor: token.colorInfoBorder, color: token.colorInfoText }">
    <Spin size="small" /><span>{{ parserMessage }}</span>
    <Button type="text" size="small" @click="emit('cancelImport')">{{ t('workspace.cancel-import') }}</Button>
  </div>
  <Alert v-if="parserStatus === 'error'" class="parser-error" type="error" :title="parserMessage" banner show-icon role="alert">
    <template #action><Button type="text" size="small" :aria-label="t('workspace.close-error')" @click="emit('dismissParserError')"><X :size="16" /></Button></template>
  </Alert>
  <Alert v-if="notice" class="notice" :type="notice.type" :title="notice.message" show-icon :role="notice.type === 'error' ? 'alert' : 'status'" :style="{ boxShadow: token.boxShadowSecondary }" />
</template>
