<script setup lang="ts">
import { reactive, watch } from 'vue'
import type { MailPropertiesDraft } from '../models/outlook'

const visible = defineModel<boolean>({ required: true })
const draft = defineModel<MailPropertiesDraft>('draft', { required: true })

defineProps<{
  outlookBusy: boolean
}>()

const editor = reactive({
  flagRequest: '',
  taskStartDate: '',
  taskDueDate: '',
})

watch(
  visible,
  (isVisible) => {
    if (!isVisible) return
    editor.flagRequest = draft.value.flagRequest
    editor.taskStartDate = draft.value.taskStartDate
    editor.taskDueDate = draft.value.taskDueDate
  },
)

function applyCustomFlag() {
  draft.value.flagRequest = editor.flagRequest
  draft.value.taskStartDate = editor.taskStartDate
  draft.value.taskDueDate = editor.taskDueDate
  visible.value = false
}
</script>

<template>
  <el-dialog v-model="visible" title="自訂旗標" width="460px" append-to-body>
    <div class="dialog-form">
      <div class="inspector-field">
        <span>旗標文字</span>
        <el-input v-model="editor.flagRequest" :disabled="outlookBusy || draft.flagInterval === 'none'" placeholder="例如：今天" />
      </div>
      <div class="date-grid">
        <div class="inspector-field">
          <span>自訂開始日</span>
          <el-date-picker
            v-model="editor.taskStartDate"
            type="date"
            value-format="YYYY-MM-DD"
            :disabled="outlookBusy || draft.flagInterval !== 'custom'"
            placeholder="選擇日期"
          />
        </div>
        <div class="inspector-field">
          <span>自訂到期日</span>
          <el-date-picker
            v-model="editor.taskDueDate"
            type="date"
            value-format="YYYY-MM-DD"
            :disabled="outlookBusy || draft.flagInterval !== 'custom'"
            placeholder="選擇日期"
          />
        </div>
      </div>
      <div v-if="draft.flagInterval !== 'custom'" class="field-hint">
        選擇「自訂日期」後即可設定自訂開始日與到期日。
      </div>
      <div class="dialog-actions">
        <el-button :disabled="outlookBusy" @click="visible = false">取消</el-button>
        <el-button type="primary" :disabled="outlookBusy" @click="applyCustomFlag">確認</el-button>
      </div>
    </div>
  </el-dialog>
</template>

<style scoped>
.dialog-form {
  display: grid;
  gap: 12px;
  min-width: 0;
}

.inspector-field {
  display: grid;
  gap: 6px;
  color: #667085;
  font-size: 0.86rem;
}

.inspector-field > span {
  font-weight: 800;
}

.date-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 8px;
}

.date-grid .el-date-editor.el-input,
.date-grid .el-date-editor.el-input__wrapper {
  width: 100%;
}

.dialog-actions {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
}

@media (max-width: 680px) {
  .date-grid {
    grid-template-columns: 1fr;
  }
}
</style>
