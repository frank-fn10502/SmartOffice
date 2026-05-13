<script setup lang="ts">
import MailPropertyPane from '../MailPropertyPane.vue'
import type { OutlookDashboardState } from '../../composables/useOutlookDashboard'
import { formatDateTime } from '../../utils/formatters'
import { formatMailSender, formatRecipient, formatRecipients, shouldShowRecipientSmtpAddress } from '../../utils/mailAddresses'

const props = defineProps<{
  dashboard: OutlookDashboardState
}>()

function formatAttachmentSize(size: number) {
  if (size >= 1024 * 1024) return `${(size / 1024 / 1024).toFixed(1)} MB`
  if (size >= 1024) return `${Math.round(size / 1024)} KB`
  return `${size} B`
}

function formatAttachmentMeta(contentType: string, size: number) {
  return `${contentType.trim() || 'unknown'} · ${formatAttachmentSize(size)}`
}

function mailPreviewText(body: string, bodyHtml: string) {
  const text = body || bodyHtml.replace(/<[^>]+>/g, ' ')
  return text.replace(/\s+/g, ' ').trim()
}

const {
  addMailCategoryDraft,
  applyMailProperties,
  categories,
  categoryTagStyle,
  closeMailDialog,
  dialogLoading,
  dialogMail,
  dialogMailAttachments,
  dialogMailConversation,
  dialogMailConversationItems,
  dialogMailFolderName,
  dialogMailHasIdentity,
  exportMailAttachment,
  flagDisplayLabel,
  flagIntervalOptions,
  flagTagType,
  isAttachmentExporting,
  isAttachmentListLoading,
  isConversationLoading,
  isMailBodyLoading,
  loadingCategories,
  mailDialogHtml,
  mailDialogVisible,
  mailHasBody,
  mailHtmlSandbox,
  mailPropertiesChanged,
  mailPropertiesDraft,
  openCategoryManager,
  openExportedAttachment,
  operationLoading,
  outlookBusy,
  removeMailCategoryDraft,
  resetMailPropertiesDraft,
  setMailFlagDraft,
  splitCategories,
} = props.dashboard
</script>

<template>
  <el-dialog
    v-model="mailDialogVisible"
    class="mail-detail-dialog"
    width="min(1160px, calc(100vw - 28px))"
    destroy-on-close
    @closed="closeMailDialog"
  >
    <template #header>
      <div v-if="dialogMail" class="dialog-mail-title">
        <strong>{{ dialogMail.subject || '(No subject)' }}</strong>
        <span>
          {{ formatMailSender(dialogMail) }} · {{ formatDateTime(dialogMail.receivedTime) }}
          <el-tag v-if="dialogLoading" size="small" type="info" effect="plain">載入中</el-tag>
        </span>
      </div>
    </template>

    <div v-if="dialogMail" class="dialog-mail-layout">
      <div class="dialog-mail-content">
        <div class="dialog-mail-summary">
          <div class="dialog-mail-meta">
            <span>寄件者</span>
            <strong>{{ formatMailSender(dialogMail) }}</strong>
            <small v-if="shouldShowRecipientSmtpAddress(dialogMail.sender)">{{ dialogMail.sender.smtpAddress }}</small>
          </div>
          <div class="dialog-mail-meta">
            <span>收件者</span>
            <strong>{{ formatRecipients(dialogMail.toRecipients) || '-' }}</strong>
            <small v-for="recipient in dialogMail.toRecipients.filter((item) => item.isGroup)" :key="recipient.displayName || recipient.smtpAddress">
              {{ formatRecipient(recipient) }} group<span v-if="recipient.members.length > 0">：{{ formatRecipients(recipient.members) }}</span>
            </small>
          </div>
          <div class="dialog-mail-meta">
            <span>Folder</span>
            <strong>{{ dialogMail.folderPath }}</strong>
          </div>
        </div>
        <div class="mail-row-tags dialog-mail-tags">
          <el-tag v-if="!dialogMail.isRead" type="warning" effect="plain">未讀</el-tag>
          <el-tag v-if="dialogMail.isMarkedAsTask" :type="flagTagType(dialogMail.flagInterval)" effect="plain">
            {{ flagDisplayLabel(dialogMail.flagInterval, dialogMail.flagRequest) }}
          </el-tag>
          <el-tag
            v-for="category in splitCategories(dialogMail.categories)"
            :key="category"
            effect="dark"
            :style="categoryTagStyle(category)"
          >
            {{ category }}
          </el-tag>
        </div>

        <div class="dialog-mail-section-head">
          <strong>內容</strong>
          <el-button v-if="mailHasBody(dialogMail)" size="small" @click="mailDialogHtml = !mailDialogHtml">
            {{ mailDialogHtml ? '切到文字' : '切到 HTML' }}
          </el-button>
        </div>
        <div class="dialog-mail-body-frame">
          <div v-if="isMailBodyLoading(dialogMail)" class="dialog-loading-block">
            <el-skeleton animated :rows="6" />
          </div>
          <iframe
            v-else-if="mailHasBody(dialogMail) && mailDialogHtml"
            class="mail-html dialog-mail-body"
            :sandbox="mailHtmlSandbox"
            referrerpolicy="no-referrer"
            :srcdoc="dialogMail.bodyHtml || dialogMail.body"
          />
          <pre v-else-if="mailHasBody(dialogMail)" class="mail-text dialog-mail-body">{{ dialogMail.body }}</pre>
          <p v-else class="hint dialog-mail-empty-body">目前沒有可顯示的 body。</p>
        </div>

        <div class="dialog-mail-attachments">
          <div class="attachment-header">
            <span class="attachment-header-title">
              <span>附件</span>
              <el-tag effect="plain">{{ dialogMailAttachments.length }}</el-tag>
            </span>
          </div>
          <div v-if="isAttachmentListLoading(dialogMail)" class="dialog-loading-block compact">
            <el-skeleton animated :rows="2" />
          </div>
          <p v-else-if="dialogMailAttachments.length === 0" class="hint">這封郵件沒有附件。</p>
          <div v-else class="attachment-list">
            <div v-for="attachment in dialogMailAttachments" :key="attachment.attachmentId" class="attachment-row">
              <span class="attachment-main">
                <strong>{{ attachment.name }}</strong>
                <span>{{ formatAttachmentMeta(attachment.contentType, attachment.size) }}</span>
              </span>
              <span class="attachment-actions">
                <el-button
                  size="small"
                  :loading="isAttachmentExporting(dialogMail, attachment)"
                  :disabled="isAttachmentExporting(dialogMail, attachment)"
                  @click="exportMailAttachment(dialogMail, attachment)"
                >
                  {{ attachment.isExported ? '重新匯出' : 'Export' }}
                </el-button>
                <el-button
                  size="small"
                  :disabled="!attachment.exportedAttachmentId"
                  @click="openExportedAttachment(attachment)"
                >
                  開啟
                </el-button>
              </span>
            </div>
          </div>
        </div>

        <div class="dialog-mail-conversation">
          <div class="attachment-header">
            <span class="attachment-header-title">
              <span>討論串</span>
              <el-tag effect="plain">{{ dialogMailConversationItems.length }}</el-tag>
            </span>
            <small v-if="dialogMailConversation?.conversationTopic">{{ dialogMailConversation.conversationTopic }}</small>
          </div>
          <div v-if="isConversationLoading(dialogMail)" class="dialog-loading-block compact">
            <el-skeleton animated :rows="2" />
          </div>
          <p v-else-if="dialogMailConversationItems.length === 0" class="hint">目前沒有可顯示的討論串。</p>
          <div v-else class="conversation-list">
            <article
              v-for="item in dialogMailConversationItems"
              :key="item.id"
              class="conversation-item"
              :class="{ current: item.id === dialogMail.id }"
            >
              <span class="conversation-marker" />
              <span class="conversation-main">
                <strong>{{ item.subject || '(No subject)' }}</strong>
                <span>{{ formatMailSender(item) }} · {{ formatDateTime(item.receivedTime) }} · {{ item.folderPath }}</span>
                <small v-if="mailPreviewText(item.body, item.bodyHtml)">{{ mailPreviewText(item.body, item.bodyHtml) }}</small>
              </span>
            </article>
          </div>
        </div>
      </div>
      <MailPropertyPane
        v-model:draft="mailPropertiesDraft"
        embedded
        :categories="categories"
        :category-tag-style="categoryTagStyle"
        :flag-interval-options="flagIntervalOptions"
        :loading-categories="loadingCategories"
        :mail-properties-changed="mailPropertiesChanged"
        :operation-loading="operationLoading"
        :outlook-busy="outlookBusy"
        :selected-mail="dialogMail"
        :selected-mail-folder-name="dialogMailFolderName"
        :selected-mail-has-identity="dialogMailHasIdentity"
        @add-category="addMailCategoryDraft"
        @apply="applyMailProperties"
        @open-category-manager="openCategoryManager"
        @remove-category="removeMailCategoryDraft"
        @reset="resetMailPropertiesDraft"
        @set-flag="setMailFlagDraft"
      />
    </div>
  </el-dialog>
</template>
