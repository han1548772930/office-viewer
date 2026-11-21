<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { Button, Input, Modal, Select } from 'antdv-next'
import { FUniver } from '../../api/facade'
import type { OfficeStudio } from '../../api/office-studio'
import type { SheetProtectionEditState, SheetProtectionViewState, SheetRangeProtectionRule } from '../../domain/model'
import { formatProtectionRange, formatProtectionUserIds } from '../../features/sheet-protection/range-protection'

const props = defineProps<{ open: boolean; studio: OfficeStudio; workbookId: string; sheetId: string; sourceRange: string }>()
const emit = defineEmits<{ 'update:open': [value: boolean]; applied: [] }>()

const editStates: Array<{ label: string; value: SheetProtectionEditState }> = [
  { label: '仅我可以编辑', value: 'onlyMe' },
  { label: '指定用户可以编辑', value: 'designedUserCanEdit' },
]
const viewStates: Array<{ label: string; value: SheetProtectionViewState }> = [
  { label: '其他用户可以查看', value: 'othersCanView' },
  { label: '仅我可以查看', value: 'noOneElseCanView' },
]
const rules = ref<SheetRangeProtectionRule[]>([])
const selectedId = ref<string>()
const description = ref('')
const editState = ref<SheetProtectionEditState>('onlyMe')
const viewState = ref<SheetProtectionViewState>('othersCanView')
const allowedUsers = ref('')
const message = ref('')
const busy = ref(false)
const selectedRule = computed(() => rules.value.find(rule => rule.id === selectedId.value))
const isOwner = computed(() => !selectedRule.value || selectedRule.value.ownerUserId === props.studio.currentUserId)

watch(() => [props.open, props.workbookId, props.sheetId] as const, ([open]) => {
  if (open) loadRules()
}, { immediate: true })

function sheet() {
  return FUniver.newAPI(props.studio).getWorkbook(props.workbookId)?.getSheetBySheetId(props.sheetId)
}

function loadRules() {
  message.value = ''
  rules.value = sheet()?.getRangeProtectionRules() ?? []
  const next = rules.value.find(rule => rule.id === selectedId.value) ?? rules.value[0]
  selectedId.value = next?.id
  loadDraft(next)
}

function loadDraft(rule?: SheetRangeProtectionRule) {
  selectedId.value = rule?.id
  description.value = rule?.description ?? ''
  editState.value = rule?.editState ?? 'onlyMe'
  viewState.value = rule?.viewState ?? 'othersCanView'
  allowedUsers.value = rule?.allowedUserIds.join(', ') ?? ''
  message.value = ''
}

function beginNewRule() {
  loadDraft()
}

function apply() {
  message.value = ''
  const activeSheet = sheet()
  if (!activeSheet) { message.value = '工作表已关闭，请重新打开范围保护。'; return }
  if (!isOwner.value) { message.value = '只有保护规则创建者可以修改该规则。'; return }
  busy.value = true
  try {
    const userIds = formatProtectionUserIds(allowedUsers.value)
    if (editState.value === 'designedUserCanEdit' && userIds.length === 0) throw new Error('请至少填写一个允许编辑的用户 ID。')
    if (selectedRule.value) {
      const rule: SheetRangeProtectionRule = {
        ...selectedRule.value,
        description: description.value.trim(),
        editState: editState.value,
        viewState: viewState.value,
        allowedUserIds: userIds,
      }
      props.studio.executeCommandSync('sheet.command.set-range-protection', {
        unitId: props.workbookId,
        sheetId: props.sheetId,
        rule,
      })
    } else {
      activeSheet.setRangeProtection(props.sourceRange, {
        ranges: [],
        description: description.value.trim(),
        editState: editState.value,
        viewState: viewState.value,
        allowedUserIds: userIds,
      })
    }
    rules.value = activeSheet.getRangeProtectionRules()
    const updated = rules.value.find(rule => rule.id === selectedRule.value?.id) ?? rules.value.at(-1)
    loadDraft(updated)
    emit('applied')
  } catch (reason) {
    message.value = reason instanceof Error ? reason.message : '范围保护规则未能保存。'
  } finally { busy.value = false }
}

function removeRule() {
  const rule = selectedRule.value
  const activeSheet = sheet()
  if (!rule || !activeSheet) return
  message.value = ''
  busy.value = true
  try {
    activeSheet.removeRangeProtection(rule.id)
    rules.value = activeSheet.getRangeProtectionRules()
    loadDraft(rules.value[0])
    emit('applied')
  } catch (reason) {
    message.value = reason instanceof Error ? reason.message : '无法移除范围保护规则。'
  } finally { busy.value = false }
}
</script>

<template>
  <Modal :open="open" title="范围保护" ok-text="保存规则" cancel-text="关闭" :confirm-loading="busy" :width="620" @ok="apply" @cancel="emit('update:open', false)">
    <div class="protection-panel">
      <aside class="protection-list" aria-label="现有保护规则">
        <Button type="text" block :class="{ 'is-selected': !selectedRule }" @click="beginNewRule">保护当前选区</Button>
        <Button
          v-for="rule in rules"
          :key="rule.id"
          type="text"
          block
          :class="{ 'is-selected': rule.id === selectedId }"
          :aria-pressed="rule.id === selectedId"
          @click="loadDraft(rule)"
        >
          <span>{{ rule.description || rule.ranges.map(formatProtectionRange).join(', ') }}</span>
        </Button>
        <p v-if="!rules.length" class="protection-empty">当前工作表没有保护规则</p>
      </aside>
      <section class="protection-form" aria-label="保护规则设置">
        <label>应用区域<Input :value="selectedRule ? selectedRule.ranges.map(formatProtectionRange).join(', ') : sourceRange" readonly aria-label="保护应用区域" /></label>
        <label>说明<Input v-model:value="description" aria-label="保护说明" :maxlength="255" :disabled="busy || !isOwner" /></label>
        <label>允许编辑<Select v-model:value="editState" :options="editStates" aria-label="保护编辑权限" :disabled="busy || !isOwner" /></label>
        <label v-if="editState === 'designedUserCanEdit'">允许编辑的用户 ID<Input v-model:value="allowedUsers" aria-label="允许编辑的用户 ID" placeholder="以逗号或空格分隔" :disabled="busy || !isOwner" /></label>
        <label>查看权限<Select v-model:value="viewState" :options="viewStates" aria-label="保护查看权限" :disabled="busy || !isOwner" /></label>
        <p v-if="selectedRule && !isOwner" class="protection-readonly">此规则由 {{ selectedRule.ownerUserId }} 创建，你只能查看。</p>
        <p v-if="message" class="protection-error" role="alert">{{ message }}</p>
        <Button v-if="selectedRule" type="text" danger :disabled="busy || !isOwner" @click="removeRule">移除所选保护规则</Button>
      </section>
    </div>
  </Modal>
</template>

<style scoped>
.protection-panel { display: grid; grid-template-columns: minmax(180px, .8fr) minmax(260px, 1.2fr); gap: 16px; min-height: 310px; }
.protection-list { display: flex; min-width: 0; flex-direction: column; align-items: stretch; gap: 4px; padding-inline-end: 10px; border-inline-end: 1px solid #e5e8ed; }
.protection-list :deep(.ant-btn) { justify-content: flex-start; min-height: 32px; height: auto; padding: 5px 8px; text-align: start; white-space: normal; }
.protection-list :deep(.ant-btn.is-selected) { background: #edf3ff; color: #2459c4; }
.protection-list :deep(.ant-btn span) { overflow-wrap: anywhere; }
.protection-empty { margin: 8px 4px; color: #687586; font-size: 12px; }
.protection-form { display: flex; min-width: 0; flex-direction: column; gap: 12px; }
.protection-form label { display: grid; min-width: 0; gap: 5px; color: #435066; font-size: 13px; }
.protection-readonly { margin: 0; color: #687586; font-size: 12px; }
.protection-error { margin: 0; color: #b42318; font-size: 12px; }
@media (max-width: 560px) { .protection-panel { grid-template-columns: minmax(0, 1fr); } .protection-list { max-height: 130px; overflow: auto; padding: 0 0 10px; border-inline-end: 0; border-bottom: 1px solid #e5e8ed; } }
</style>
