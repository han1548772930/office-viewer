<script setup lang="ts">
import { computed, inject, ref, watch } from 'vue'
import { Button, Form, FormItem, Input, Modal, Select } from 'antdv-next'
import type { DocumentPermissionRule, DocumentPermissionScope } from '../../domain/model.ts'
import { validateDocumentPermissionRules } from '../../features/document/permissions/document-permissions.ts'
import { createDocumentPermissionRule, removeDocumentPermissionRule, upsertDocumentPermissionRule, type DocumentPermissionTargetOption } from '../../features/document/permissions/permission-rules.ts'
import { OfficeTranslatorKey } from '../../i18n/office-i18n.ts'

const props = withDefaults(defineProps<{
  open: boolean
  rules: readonly DocumentPermissionRule[]
  targets: readonly DocumentPermissionTargetOption[]
  readOnly?: boolean
}>(), { readOnly: false })

const emit = defineEmits<{
  'update:open': [value: boolean]
  save: [rules: DocumentPermissionRule[]]
}>()

const t = inject(OfficeTranslatorKey, (_key: string, fallback: string) => fallback)
const rules = ref<DocumentPermissionRule[]>([])
const selectedId = ref<string>()
const scope = ref<DocumentPermissionScope>('document')
const targetId = ref('')
const action = ref<DocumentPermissionRule['action']>('edit')
const effect = ref<DocumentPermissionRule['effect']>('deny')
const userIds = ref('')
const error = ref('')
const selectedRule = computed(() => rules.value.find(rule => rule.id === selectedId.value))
const scopedTargets = computed(() => props.targets.filter(target => target.scope === scope.value))
const actionOptions = computed(() => scope.value === 'document'
  ? [
      { value: 'edit', label: t('doc.permissions.action.edit', 'Edit') },
      { value: 'copy', label: t('doc.permissions.action.copy', 'Copy') },
      { value: 'comment', label: t('doc.permissions.action.comment', 'Comment') },
      { value: 'print', label: t('doc.permissions.action.print', 'Print') },
      { value: 'export', label: t('doc.permissions.action.export', 'Export') },
    ]
  : [{ value: 'edit', label: t('doc.permissions.action.edit', 'Edit') }])

watch(() => [props.open, props.rules] as const, ([open, incoming]) => {
  if (!open) return
  rules.value = incoming.map(rule => ({ ...rule, userIds: [...rule.userIds] }))
  const rule = rules.value.find(candidate => candidate.id === selectedId.value) ?? rules.value[0]
  loadRule(rule)
  error.value = ''
}, { immediate: true })

function loadRule(rule?: DocumentPermissionRule) {
  selectedId.value = rule?.id
  scope.value = rule?.scope ?? 'document'
  targetId.value = rule?.targetId ?? ''
  action.value = rule?.action ?? 'edit'
  effect.value = rule?.effect ?? 'deny'
  userIds.value = rule?.userIds.join(', ') ?? ''
  error.value = ''
}

function beginNewRule() {
  loadRule()
}

function changeScope(value: DocumentPermissionScope) {
  scope.value = value
  targetId.value = ''
  if (value !== 'document') action.value = 'edit'
}

function parseUserIds(value: string): string[] {
  return [...new Set(value.split(/[\s,;]+/).map(userId => userId.trim()).filter(Boolean))]
}

function saveRule() {
  error.value = ''
  try {
    const draft = {
      scope: scope.value,
      ...(scope.value === 'document' ? {} : { targetId: targetId.value }),
      action: action.value,
      effect: effect.value,
      userIds: parseUserIds(userIds.value),
    }
    const current = selectedRule.value
    const rule = current
      ? { ...draft, id: current.id } as DocumentPermissionRule
      : createDocumentPermissionRule(rules.value, draft)
    const next = upsertDocumentPermissionRule(rules.value, rule)
    validateDocumentPermissionRules(next)
    rules.value = next
    selectedId.value = rule.id
    emit('save', next.map(item => ({ ...item, userIds: [...item.userIds] })))
  } catch (reason) {
    error.value = reason instanceof Error ? reason.message : t('doc.permissions.save-failed', 'Could not save the permission rule.')
  }
}

function deleteRule() {
  if (!selectedRule.value) return
  try {
    rules.value = removeDocumentPermissionRule(rules.value, selectedRule.value.id)
    loadRule(rules.value[0])
    emit('save', rules.value.map(item => ({ ...item, userIds: [...item.userIds] })))
  } catch (reason) {
    error.value = reason instanceof Error ? reason.message : t('doc.permissions.delete-failed', 'Could not delete the permission rule.')
  }
}

function scopeName(value: DocumentPermissionScope): string {
  return t(`doc.permissions.scope.${value}`, value === 'document' ? 'Document' : value === 'section' ? 'Section' : value === 'paragraph' ? 'Paragraph' : 'Object')
}

function ruleLabel(rule: DocumentPermissionRule): string {
  const targetLabel = rule.scope === 'document'
    ? ''
    : props.targets.find(target => target.scope === rule.scope && target.id === rule.targetId)?.label ?? rule.targetId
  const users = rule.userIds.length ? rule.userIds.join(', ') : t('doc.permissions.all-users', 'All users')
  return [scopeName(rule.scope), targetLabel, rule.effect === 'deny' ? t('doc.permissions.deny', 'Deny') : t('doc.permissions.allow', 'Allow'), users].filter(Boolean).join(' · ')
}
</script>

<template>
  <Modal :open="open" :title="t('doc.permissions.title', 'Document permissions')" :cancel-text="t('common.close', 'Close')" :footer="null" :width="700" @cancel="emit('update:open', false)">
    <div class="permission-panel">
      <aside class="permission-list" :aria-label="t('doc.permissions.rules', 'Permission rules')">
        <Button type="text" block :disabled="readOnly" :class="{ 'is-selected': !selectedRule }" @click="beginNewRule">{{ t('doc.permissions.add-rule', 'Add rule') }}</Button>
        <Button v-for="rule in rules" :key="rule.id" type="text" block :aria-pressed="rule.id === selectedId" :class="{ 'is-selected': rule.id === selectedId }" @click="loadRule(rule)">
          <span>{{ ruleLabel(rule) }}</span>
        </Button>
        <p v-if="!rules.length" class="permission-empty">{{ t('doc.permissions.empty', 'No permission rules') }}</p>
      </aside>
      <Form layout="vertical" class="permission-form" :aria-label="t('doc.permissions.rule-settings', 'Rule settings')">
        <FormItem :label="t('doc.permissions.scope', 'Scope')">
          <Select :value="scope" :options="[
            { value: 'document', label: scopeName('document') },
            { value: 'section', label: scopeName('section') },
            { value: 'paragraph', label: scopeName('paragraph') },
            { value: 'entity', label: scopeName('entity') },
          ]" :disabled="readOnly" @update:value="changeScope($event as DocumentPermissionScope)" />
        </FormItem>
        <FormItem v-if="scope !== 'document'" :label="t('doc.permissions.target', 'Target')">
          <Select v-model:value="targetId" show-search :options="scopedTargets.map(target => ({ value: target.id, label: target.label }))" :disabled="readOnly" :placeholder="t('doc.permissions.choose-target', 'Choose a target')" />
        </FormItem>
        <FormItem :label="t('doc.permissions.action', 'Action')">
          <Select v-model:value="action" :options="actionOptions" :disabled="readOnly" />
        </FormItem>
        <FormItem :label="t('doc.permissions.effect', 'Effect')">
          <Select v-model:value="effect" :options="[
            { value: 'allow', label: t('doc.permissions.allow', 'Allow') },
            { value: 'deny', label: t('doc.permissions.deny', 'Deny') },
          ]" :disabled="readOnly" />
        </FormItem>
        <FormItem :label="t('doc.permissions.users', 'User IDs')">
          <Input.TextArea v-model:value="userIds" :disabled="readOnly" :placeholder="t('doc.permissions.users-placeholder', 'Leave empty to apply to all users')" :auto-size="{ minRows: 2, maxRows: 4 }" />
        </FormItem>
        <p v-if="error" class="permission-error" role="alert">{{ error }}</p>
        <div class="permission-actions">
          <Button type="primary" :disabled="readOnly || (scope !== 'document' && !targetId)" @click="saveRule">{{ t('doc.permissions.save-rule', 'Save rule') }}</Button>
          <Button v-if="selectedRule" danger :disabled="readOnly" @click="deleteRule">{{ t('doc.permissions.delete-rule', 'Delete rule') }}</Button>
        </div>
      </Form>
    </div>
  </Modal>
</template>

<style scoped>
.permission-panel { display: grid; grid-template-columns: minmax(220px, .85fr) minmax(280px, 1.15fr); gap: 16px; min-height: 320px; }
.permission-list { display: flex; min-width: 0; flex-direction: column; align-items: stretch; gap: 4px; max-height: 430px; overflow: auto; padding-inline-end: 10px; border-inline-end: 1px solid #e5e8ed; }
.permission-list :deep(.ant-btn) { justify-content: flex-start; min-height: 34px; height: auto; padding: 6px 8px; text-align: start; white-space: normal; }
.permission-list :deep(.ant-btn.is-selected) { background: #edf3ff; color: #2459c4; }
.permission-list :deep(.ant-btn span) { overflow-wrap: anywhere; }
.permission-empty { margin: 8px 4px; color: #687586; font-size: 12px; }
.permission-form { display: flex; min-width: 0; flex-direction: column; gap: 12px; }
.permission-form :deep(.ant-form-item) { margin-bottom: 0; }
.permission-error { margin: 0; color: #b42318; font-size: 12px; overflow-wrap: anywhere; }
.permission-actions { display: flex; gap: 8px; }
@media (max-width: 600px) { .permission-panel { grid-template-columns: minmax(0, 1fr); } .permission-list { max-height: 150px; padding: 0 0 10px; border-inline-end: 0; border-bottom: 1px solid #e5e8ed; } }
</style>
