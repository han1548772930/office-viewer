<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { Alert, Button, Form, FormItem, Input, Modal, Select } from 'antdv-next'
import type { OfficeStudio } from '../../api/office-studio'
import type { SheetWorksheetProtectionRule } from '../../domain/model'
import { createWorksheetProtectionRule } from '../../features/sheet-protection/range-protection'

const props = defineProps<{
  open: boolean
  studio: OfficeStudio
  workbookId: string
  sheetId: string
  rule?: SheetWorksheetProtectionRule
}>()
const emit = defineEmits<{ 'update:open': [value: boolean]; applied: [] }>()

const description = ref('')
const editState = ref<'onlyMe' | 'designedUserCanEdit'>('onlyMe')
const viewState = ref<'othersCanView' | 'noOneElseCanView'>('othersCanView')
const allowedUsers = ref('')
const error = ref('')
const busy = ref(false)
const editingExisting = computed(() => Boolean(props.rule))
const canManage = computed(() => !props.rule || props.rule.ownerUserId === props.studio.currentUserId)
const editOptions = [
  { label: '仅所有者可以编辑', value: 'onlyMe' },
  { label: '指定用户可以编辑', value: 'designedUserCanEdit' },
]
const viewOptions = [
  { label: '其他用户可以查看', value: 'othersCanView' },
  { label: '仅所有者可以查看', value: 'noOneElseCanView' },
]

watch(() => [props.open, props.rule, props.sheetId] as const, ([open, rule]) => {
  if (!open) return
  description.value = rule?.description ?? ''
  editState.value = rule?.editState ?? 'onlyMe'
  viewState.value = rule?.viewState ?? 'othersCanView'
  allowedUsers.value = rule?.allowedUserIds.join(', ') ?? ''
  error.value = ''
}, { immediate: true })

function apply() {
  error.value = ''
  if (!canManage.value) return
  busy.value = true
  try {
    const allowedUserIds = [...new Set(allowedUsers.value.split(/[\s,;]+/).map(value => value.trim()).filter(Boolean))]
    const input = { description: description.value, editState: editState.value, viewState: viewState.value, allowedUserIds }
    const created = createWorksheetProtectionRule(input, props.studio.currentUserId)
    const rule = props.rule ? { ...created, id: props.rule.id, permissionId: props.rule.permissionId } : created
    if (props.rule && props.rule.ownerUserId !== props.studio.currentUserId) throw new Error('只有工作表保护创建者可以修改该保护。')
    props.studio.executeCommandSync('sheet.command.add-worksheet-protection', {
      unitId: props.workbookId,
      sheetId: props.sheetId,
      rule: { ...rule, ownerUserId: props.studio.currentUserId, unitType: 'worksheet' },
    })
    emit('applied')
    emit('update:open', false)
  } catch (reason) {
    error.value = reason instanceof Error ? reason.message : '无法保存工作表保护设置。'
  } finally {
    busy.value = false
  }
}

function remove() {
  if (!props.rule || !canManage.value) return
  error.value = ''
  busy.value = true
  try {
    props.studio.executeCommandSync('sheet.command.delete-worksheet-protection', {
      unitId: props.workbookId,
      sheetId: props.sheetId,
    })
    emit('applied')
    emit('update:open', false)
  } catch (reason) {
    error.value = reason instanceof Error ? reason.message : '无法移除工作表保护。'
  } finally {
    busy.value = false
  }
}
</script>

<template>
  <Modal :open="open" :title="editingExisting ? '工作表保护设置' : '保护工作表'" :ok-text="editingExisting ? '保存设置' : '启用保护'" :ok-button-props="{ disabled: !canManage }" cancel-text="取消" :confirm-loading="busy" :width="420" @ok="apply" @cancel="emit('update:open', false)">
    <Form layout="vertical" class="worksheet-protection-form" @finish="apply">
      <FormItem label="说明">
        <Input v-model:value="description" :maxlength="255" :disabled="busy || !canManage" aria-label="工作表保护说明" />
      </FormItem>
      <FormItem label="允许编辑">
        <Select v-model:value="editState" :options="editOptions" :disabled="busy || !canManage" aria-label="工作表编辑权限" />
      </FormItem>
      <FormItem v-if="editState === 'designedUserCanEdit'" label="允许编辑的用户 ID">
        <Input v-model:value="allowedUsers" placeholder="以逗号或空格分隔" :disabled="busy || !canManage" aria-label="允许编辑的用户 ID" />
      </FormItem>
      <FormItem label="查看权限">
        <Select v-model:value="viewState" :options="viewOptions" :disabled="busy || !canManage" aria-label="工作表查看权限" />
      </FormItem>
      <Alert v-if="rule && !canManage" type="info" :message="`该工作表由 ${rule.ownerUserId} 保护，你可以查看设置但不能修改。`" />
      <Alert v-if="error" type="error" :message="error" role="alert" />
      <Button v-if="editingExisting" type="text" danger :disabled="busy || !canManage" @click="remove">移除工作表保护</Button>
    </Form>
  </Modal>
</template>

<style scoped>
.worksheet-protection-form { display: flex; flex-direction: column; gap: 10px; }
.worksheet-protection-form :deep(.ant-form-item) { margin-bottom: 0; }
</style>
