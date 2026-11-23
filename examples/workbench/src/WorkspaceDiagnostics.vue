<script setup lang="ts">
import { Collapse, theme } from 'antdv-next'
import { computed } from 'vue'
import { useI18n } from 'vue-i18n'
import type { ParserDiagnostic } from '@office-studio/office-studio'

defineProps<{ diagnostics: readonly ParserDiagnostic[] }>()
const { t } = useI18n()
const { token } = theme.useToken()
const diagnosticsTheme = computed(() => ({
  '--workspace-diagnostic-text': token.value.colorText,
  '--workspace-diagnostic-muted': token.value.colorTextSecondary,
  '--workspace-diagnostic-border': token.value.colorBorderSecondary,
  '--workspace-diagnostic-error': token.value.colorError,
  '--workspace-diagnostic-warning': token.value.colorWarning,
}))
</script>

<template>
  <div :style="diagnosticsTheme">
    <p v-if="!diagnostics.length" class="workspace-diagnostics-empty">{{ t('workspace.no-diagnostics') }}</p>
    <ul v-else class="workspace-diagnostics">
      <li v-for="(diagnostic, index) in diagnostics" :key="diagnostic.code + ':' + index" :data-severity="diagnostic.severity">
        <strong>{{ t(`workspace.severity-${diagnostic.severity}`) }} · {{ diagnostic.code }}</strong>
        <p>{{ diagnostic.message }}</p>
        <Collapse v-if="diagnostic.details" :items="[{ key: 'details', label: t('workspace.details') }]" ghost size="small" destroy-on-hidden>
          <template #contentRender><pre>{{ diagnostic.details }}</pre></template>
        </Collapse>
      </li>
    </ul>
  </div>
</template>
