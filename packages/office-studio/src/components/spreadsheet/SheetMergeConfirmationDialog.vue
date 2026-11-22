<script setup lang="ts">
import { Alert, Modal } from 'antdv-next'
import type { SheetMergeConfirmationRequest } from '../../features/sheet-workbench/sheet-merge-command.ts'

defineProps<{ request: Pick<SheetMergeConfirmationRequest, 'discardedContentCount'> | null; error: string }>()
const emit = defineEmits<{ confirm: []; cancel: []; closed: [] }>()
</script>

<template>
  <Modal :open="Boolean(request)" title="合并单元格" ok-text="继续合并" cancel-text="取消" :width="440"
    :mask-closable="false" :focus-trigger-after-close="false" @ok="emit('confirm')" @cancel="emit('cancel')" @after-close="emit('closed')">
    <p>每个合并区域保留按行列顺序找到的第一项非空内容，其余 {{ request?.discardedContentCount ?? 0 }} 项内容将被清除。</p>
    <Alert v-if="error" type="error" :title="error" show-icon role="alert" />
  </Modal>
</template>
