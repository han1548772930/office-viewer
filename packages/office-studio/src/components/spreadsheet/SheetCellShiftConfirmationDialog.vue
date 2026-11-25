<script setup lang="ts">
import { Alert, Modal } from 'antdv-next'
import type { SheetCellShiftConfirmationRequest } from '../../features/sheet-workbench/sheet-cell-shift-command.ts'
import { formatProtectionRange } from '../../features/sheet-protection/range-protection.ts'
import { normalizeSheetCommandRange } from '../../features/sheet-workbench/sheet-range-command.ts'

defineProps<{ request: SheetCellShiftConfirmationRequest | null; error: string }>()
const emit = defineEmits<{ confirm: []; cancel: []; closed: [] }>()
</script>

<template>
  <Modal :open="Boolean(request)" title="合并单元格" ok-text="继续" cancel-text="取消" :width="440"
    :mask-closable="false" :focus-trigger-after-close="false" @ok="emit('confirm')" @cancel="emit('cancel')" @after-close="emit('closed')">
    <p>移动 {{ request?.input.range ? formatProtectionRange(normalizeSheetCommandRange(request.input.range)) : '' }} 时，需要拆分以下 {{ request?.mergeRanges.length ?? 0 }} 个合并区域：</p>
    <p v-for="range in request?.mergeRanges.slice(0, 8)" :key="formatProtectionRange(range)">{{ formatProtectionRange(range) }}</p>
    <p v-if="request && request.mergeRanges.length > 8">其余 {{ request.mergeRanges.length - 8 }} 个区域也会拆分。</p>
    <Alert v-if="error" type="error" :title="error" show-icon role="alert" />
  </Modal>
</template>
