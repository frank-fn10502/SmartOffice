<script setup lang="ts">
import { PriceTag } from '@element-plus/icons-vue'
import type { MailItemDto, MailPropertiesDraft, OutlookCategoryDto } from '../models/outlook'
import { formatDateTime } from '../utils/formatters'
import { formatMailSender, shouldShowRecipientSmtpAddress } from '../utils/mailAddresses'
import { canUpdateMailProperties, outlookItemTypeLabel } from '../utils/outlookItemTypes'
import { flagDisplayLabel, flagTagType } from '../utils/outlookDashboardHelpers'

const draft = defineModel<MailPropertiesDraft>('draft', { required: true })

defineProps<{
  embedded?: boolean
  selectedMail: MailItemDto | null
  selectedMailFolderName: string
  selectedMailHasIdentity: boolean
  categories: OutlookCategoryDto[]
  loadingCategories: boolean
  flagIntervalOptions: { label: string; value: string }[]
  mailPropertiesChanged: boolean
  operationLoading: boolean
  outlookBusy: boolean
  categoryTagStyle: (name: string) => Record<string, string>
}>()

defineEmits<{
  addCategory: [name: string]
  apply: [mail: MailItemDto]
  openCategoryManager: []
  removeCategory: [name: string]
  reset: [mail: MailItemDto]
  setFlag: [interval: string]
}>()
</script>

<template>
  <section :class="embedded ? 'outlook-property-pane embedded-property-pane' : 'panel outlook-property-pane'">
    <div class="panel-header">
      <div class="panel-title">
        <el-icon><PriceTag /></el-icon>
        <span>修改郵件屬性</span>
      </div>
    </div>

    <div class="inspector-panel-body">
      <div v-if="selectedMail" class="mail-inspector">
        <div class="inspector-subject">{{ selectedMail.subject }}</div>
        <div class="inspector-meta">
          <span>
            {{ formatMailSender(selectedMail) }}<template v-if="shouldShowRecipientSmtpAddress(selectedMail.sender)"> &lt;{{ selectedMail.sender.smtpAddress }}&gt;</template>
          </span>
          <span>{{ formatDateTime(selectedMail.receivedTime) }}</span>
          <span>來源：{{ selectedMailFolderName }}</span>
        </div>
        <div v-if="!selectedMailHasIdentity" class="identity-warning">
          這封郵件缺少 id，Add-in 需在 PushMails 回傳 Outlook EntryID 或穩定識別後才能修改或移動。
        </div>
        <div v-else-if="!canUpdateMailProperties(selectedMail)" class="identity-warning">
          這是 {{ outlookItemTypeLabel(selectedMail) }}，目前可閱讀內容、附件，也可拖曳移動；分類、旗標與已讀狀態仍需使用 Outlook 的會議/行事曆流程。
        </div>

        <div class="mail-property-form">
          <div class="inspector-field">
            <span>已讀/未讀</span>
            <div class="property-tag-picker">
              <el-tag
                class="clickable-marker-tag"
                :class="{ disabled: outlookBusy || !canUpdateMailProperties(selectedMail) }"
                :type="draft.isRead ? 'info' : 'warning'"
                effect="plain"
                role="button"
                tabindex="0"
                :aria-disabled="outlookBusy || !canUpdateMailProperties(selectedMail)"
                @click="!outlookBusy && canUpdateMailProperties(selectedMail) && (draft.isRead = !draft.isRead)"
                @keydown.enter.prevent="!outlookBusy && canUpdateMailProperties(selectedMail) && (draft.isRead = !draft.isRead)"
                @keydown.space.prevent="!outlookBusy && canUpdateMailProperties(selectedMail) && (draft.isRead = !draft.isRead)"
              >
                {{ draft.isRead ? '已讀' : '未讀' }}
              </el-tag>
            </div>
          </div>

          <div class="inspector-field">
            <span class="field-title-row">
              <span>Category</span>
              <el-button size="small" text @click="$emit('openCategoryManager')">
                管理分類 ({{ categories.length }})
              </el-button>
            </span>
            <div class="property-tag-picker">
              <el-tag
                v-for="category in draft.categories"
                :key="category"
                class="property-removable-tag"
                :closable="canUpdateMailProperties(selectedMail)"
                effect="dark"
                :disable-transitions="true"
                :style="categoryTagStyle(category)"
                @close="canUpdateMailProperties(selectedMail) && $emit('removeCategory', category)"
              >
                {{ category }}
              </el-tag>
              <span v-if="draft.categories.length === 0" class="field-hint">尚未套用分類。</span>
            </div>
            <div class="property-tag-picker">
              <el-tag
                v-for="category in categories.filter((item) => !draft.categories.some((selected) => selected.toLowerCase() === item.name.toLowerCase()))"
                :key="category.name"
                class="clickable-marker-tag"
                :class="{ disabled: outlookBusy || !canUpdateMailProperties(selectedMail) }"
                type="info"
                effect="plain"
                role="button"
                tabindex="0"
                :aria-disabled="outlookBusy || !canUpdateMailProperties(selectedMail)"
                @click="canUpdateMailProperties(selectedMail) && $emit('addCategory', category.name)"
                @keydown.enter.prevent="canUpdateMailProperties(selectedMail) && $emit('addCategory', category.name)"
                @keydown.space.prevent="canUpdateMailProperties(selectedMail) && $emit('addCategory', category.name)"
              >
                {{ category.name }}
              </el-tag>
              <span v-if="categories.length === 0 && !loadingCategories" class="field-hint">尚未取得可套用的分類。</span>
            </div>
          </div>

          <div class="inspector-field">
            <span>Flag</span>
            <div class="property-tag-picker">
              <el-tag
                v-for="option in flagIntervalOptions"
                :key="option.value"
                class="clickable-marker-tag"
                :class="{ disabled: outlookBusy || !canUpdateMailProperties(selectedMail) }"
                :type="flagTagType(option.value, draft.flagInterval === option.value)"
                :effect="draft.flagInterval === option.value ? 'dark' : 'plain'"
                role="button"
                tabindex="0"
                :aria-pressed="draft.flagInterval === option.value"
                :aria-disabled="outlookBusy || !canUpdateMailProperties(selectedMail)"
                @click="canUpdateMailProperties(selectedMail) && $emit('setFlag', option.value)"
                @keydown.enter.prevent="canUpdateMailProperties(selectedMail) && $emit('setFlag', option.value)"
                @keydown.space.prevent="canUpdateMailProperties(selectedMail) && $emit('setFlag', option.value)"
              >
                {{ option.label }}
              </el-tag>
            </div>
            <div
              v-if="draft.flagInterval !== 'none'"
              class="flag-draft-summary"
              :class="{ complete: draft.flagInterval === 'complete' }"
            >
              <span class="flag-draft-label">{{ flagDisplayLabel(draft.flagInterval, draft.flagRequest) }}</span>
              <span v-if="draft.taskDueDate" class="flag-draft-date">到期 {{ draft.taskDueDate }}</span>
              <span v-else-if="draft.flagInterval === 'custom'" class="flag-draft-date muted">到期 未設定</span>
              <span v-if="draft.flagInterval === 'custom' && draft.taskStartDate" class="flag-draft-date secondary">
                開始 {{ draft.taskStartDate }}
              </span>
              <span v-if="draft.flagInterval === 'complete' && draft.taskCompletedDate" class="flag-draft-date secondary">
                完成 {{ draft.taskCompletedDate }}
              </span>
            </div>
          </div>

          <div class="inspector-actions commit-actions">
            <el-button @click="$emit('reset', selectedMail)">重設</el-button>
            <el-button
              type="primary"
              size="large"
              class="commit-button"
              :loading="operationLoading"
              :disabled="!selectedMailHasIdentity || !canUpdateMailProperties(selectedMail) || !mailPropertiesChanged || (outlookBusy && !operationLoading)"
              @click="$emit('apply', selectedMail)"
            >
              送出並更新 Outlook
            </el-button>
          </div>
        </div>

        <div class="inspector-note">郵件屬性會一次送到 Outlook；移動郵件請拖曳中央郵件到左側 folder。等待 Add-in 回推前會鎖住操作。</div>
      </div>

      <div v-else class="empty-inspector">
        請先選取中央郵件，這裡會顯示該郵件目前的屬性與可修改欄位。
      </div>
    </div>
  </section>
</template>

<style scoped>
.embedded-property-pane {
  min-width: 0;
  min-height: 0;
  height: 100%;
  overflow: hidden;
  border-left: 1px solid #edf0f5;
  padding-left: 16px;
}

.embedded-property-pane .panel-header {
  min-height: 38px;
  padding: 0 0 10px;
  border-bottom: 0;
  background: transparent;
}

.embedded-property-pane .inspector-panel-body {
  min-height: 0;
  overflow: auto;
  padding: 0;
}

.mail-inspector {
  display: grid;
  gap: 12px;
  min-width: 0;
}

.inspector-subject {
  min-width: 0;
  color: #172033;
  font-size: 1.1rem;
  font-weight: 800;
  line-height: 1.35;
  overflow-wrap: anywhere;
}

.inspector-meta,
.inspector-field {
  display: grid;
  gap: 6px;
  color: #667085;
  font-size: 0.86rem;
}

.inspector-meta > span,
.inspector-field > span {
  min-width: 0;
  overflow-wrap: anywhere;
}

.inspector-field > span {
  font-weight: 800;
}

.identity-warning {
  padding: 9px 10px;
  border: 1px solid #f7c948;
  border-radius: 6px;
  background: #fff8db;
  color: #8a5a00;
  font-size: 0.84rem;
  line-height: 1.45;
}

.mail-property-form {
  display: grid;
  gap: 10px;
  min-width: 0;
  padding: 2px 0 4px;
}

.field-title-row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  flex-wrap: wrap;
  gap: 8px;
  min-width: 0;
}

.field-title-row .el-button {
  min-width: 0;
  margin-left: auto;
  white-space: normal;
}

.property-tag-picker {
  display: flex;
  min-width: 0;
  flex-wrap: wrap;
  gap: 8px;
}

.property-tag-picker .el-tag {
  max-width: 100%;
}

.property-tag-picker :deep(.el-tag__content) {
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
}

.property-removable-tag :deep(.el-tag__content),
.clickable-marker-tag :deep(.el-tag__content) {
  overflow-wrap: anywhere;
  white-space: normal;
}

.flag-draft-summary {
  display: flex;
  min-width: 0;
  flex-wrap: wrap;
  align-items: center;
  gap: 8px;
  padding: 8px 10px;
  border: 1px solid #f7c7c7;
  border-radius: 6px;
  background: #fff7f7;
  color: #912018;
  font-size: 0.84rem;
  line-height: 1.45;
}

.flag-draft-summary.complete {
  border-color: #a8d5ba;
  background: #f1fbf5;
  color: #176b3a;
}

.flag-draft-summary > span {
  max-width: 100%;
  overflow-wrap: anywhere;
}

.flag-draft-label {
  font-weight: 800;
}

.flag-draft-date {
  color: #b42318;
  font-weight: 700;
}

.flag-draft-summary.complete .flag-draft-date {
  color: #087443;
}

.flag-draft-date.secondary {
  color: #475467;
  font-weight: 600;
}

.flag-draft-date.muted {
  color: #b42318;
  font-weight: 600;
}

.property-removable-tag {
  --el-tag-text-color: inherit;
}

.clickable-marker-tag {
  cursor: pointer;
  user-select: none;
}

.clickable-marker-tag.disabled {
  cursor: not-allowed;
  opacity: 0.65;
}

.clickable-marker-tag:focus-visible {
  outline: 2px solid #1f5f99;
  outline-offset: 2px;
}

.inspector-actions {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
}

.inspector-note {
  padding: 9px 10px;
  border-radius: 6px;
  background: #f8fafc;
  color: #667085;
  font-size: 0.82rem;
  line-height: 1.45;
}

.commit-actions {
  display: grid;
  grid-template-columns: 1fr;
  gap: 8px;
  padding-top: 2px;
}

.commit-actions .el-button {
  width: 100%;
  min-height: 44px;
}

.commit-actions .el-button + .el-button {
  margin-left: 0;
}

.commit-button {
  font-weight: 800;
}

.empty-inspector {
  display: grid;
  min-height: 200px;
  align-items: center;
  color: #667085;
  font-size: 0.9rem;
  line-height: 1.5;
}
</style>
