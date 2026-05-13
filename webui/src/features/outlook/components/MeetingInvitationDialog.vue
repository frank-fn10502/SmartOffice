<script setup lang="ts">
import { ElMessage } from 'element-plus'
import type { OutlookDashboardState } from '../composables/useOutlookDashboard'
import { formatDateTime } from '../utils/formatters'
import { formatMailSender, formatRecipient, formatRecipients, shouldShowRecipientSmtpAddress } from '../utils/mailAddresses'
import { outlookItemTypeLabel } from '../utils/outlookItemTypes'

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

const {
  closeMailDialog,
  activeView,
  calendarEvents,
  dialogLoading,
  dialogMail,
  dialogMailAttachments,
  exportMailAttachment,
  isAttachmentExporting,
  isAttachmentListLoading,
  isMailBodyLoading,
  mailDialogHtml,
  mailDialogVisible,
  mailHasBody,
  mailHtmlSandbox,
  openExportedAttachment,
  requestCalendar,
  selectCalendarEvent,
  splitCategories,
} = props.dashboard

function normalizeSubject(value: string) {
  return value
    .toLowerCase()
    .replace(/^re:\s*/i, '')
    .replace(/^fw:\s*/i, '')
    .replace(/^fwd:\s*/i, '')
    .replace(/^會議邀請[:：]\s*/i, '')
    .replace(/^會議更新[:：]\s*/i, '')
    .replace(/\s+/g, '')
}

function findMatchingCalendarEvent(subject: string) {
  const normalizedSubject = normalizeSubject(subject)
  if (!normalizedSubject) return null
  return calendarEvents.value.find((event) => {
    const normalizedEventSubject = normalizeSubject(event.subject)
    return normalizedEventSubject === normalizedSubject
      || normalizedEventSubject.includes(normalizedSubject)
      || normalizedSubject.includes(normalizedEventSubject)
  }) ?? null
}

async function openCalendarView() {
  activeView.value = 'calendar'
  mailDialogVisible.value = false
  await requestCalendar()
  const event = dialogMail.value ? findMatchingCalendarEvent(dialogMail.value.subject) : null
  if (event) {
    selectCalendarEvent(event)
    return
  }
  ElMessage.warning('已切到月曆，但目前沒有找到可匹配的 calendar event。')
}
</script>

<template>
  <el-dialog
    v-model="mailDialogVisible"
    class="meeting-invitation-dialog"
    width="min(1040px, calc(100vw - 28px))"
    destroy-on-close
    @closed="closeMailDialog"
  >
    <template #header>
      <div v-if="dialogMail" class="meeting-dialog-title">
        <span class="meeting-kind">{{ outlookItemTypeLabel(dialogMail) || '會議邀請' }}</span>
        <strong>{{ dialogMail.subject || '(No subject)' }}</strong>
        <span>
          {{ formatMailSender(dialogMail) }} · {{ formatDateTime(dialogMail.receivedTime) }}
          <el-tag v-if="dialogLoading" size="small" type="info" effect="plain">載入中</el-tag>
        </span>
      </div>
    </template>

    <div v-if="dialogMail" class="meeting-dialog-layout">
      <aside class="meeting-dialog-side">
        <section class="meeting-info-panel">
          <h3>邀請資訊</h3>
          <dl>
            <div>
              <dt>邀請者</dt>
              <dd>
                <strong>{{ formatMailSender(dialogMail) }}</strong>
                <small v-if="shouldShowRecipientSmtpAddress(dialogMail.sender)">{{ dialogMail.sender.smtpAddress }}</small>
              </dd>
            </div>
            <div>
              <dt>與會者</dt>
              <dd>
                <strong>{{ formatRecipients(dialogMail.toRecipients) || '-' }}</strong>
                <small v-for="recipient in dialogMail.toRecipients.filter((item) => item.isGroup)" :key="recipient.displayName || recipient.smtpAddress">
                  {{ formatRecipient(recipient) }} group<span v-if="recipient.members.length > 0">：{{ formatRecipients(recipient.members) }}</span>
                </small>
              </dd>
            </div>
            <div v-if="dialogMail.ccRecipients.length > 0">
              <dt>副本</dt>
              <dd>
                <strong>{{ formatRecipients(dialogMail.ccRecipients) }}</strong>
                <small v-for="recipient in dialogMail.ccRecipients.filter((item) => item.isGroup)" :key="recipient.displayName || recipient.smtpAddress">
                  {{ formatRecipient(recipient) }} group<span v-if="recipient.members.length > 0">：{{ formatRecipients(recipient.members) }}</span>
                </small>
              </dd>
            </div>
            <div>
              <dt>Folder</dt>
              <dd><strong>{{ dialogMail.folderPath }}</strong></dd>
            </div>
          </dl>
        </section>

        <section class="meeting-info-panel">
          <h3>狀態</h3>
          <div class="meeting-tags">
            <el-tag type="success" effect="plain">{{ dialogMail.messageClass || 'IPM.Schedule.Meeting' }}</el-tag>
            <el-tag v-if="!dialogMail.isRead" type="warning" effect="plain">未讀</el-tag>
            <el-tag
              v-for="category in splitCategories(dialogMail.categories)"
              :key="category"
              effect="plain"
            >
              {{ category }}
            </el-tag>
          </div>
        </section>
      </aside>

      <main class="meeting-dialog-main">
        <div class="meeting-section-head">
          <strong>會議內容</strong>
          <span class="meeting-section-actions">
            <el-button size="small" type="primary" plain @click="openCalendarView">
              前往月曆
            </el-button>
            <el-button v-if="mailHasBody(dialogMail)" size="small" @click="mailDialogHtml = !mailDialogHtml">
              {{ mailDialogHtml ? '切到文字' : '切到 HTML' }}
            </el-button>
          </span>
        </div>

        <div class="meeting-body-frame">
          <div v-if="isMailBodyLoading(dialogMail)" class="dialog-loading-block">
            <el-skeleton animated :rows="7" />
          </div>
          <iframe
            v-else-if="mailHasBody(dialogMail) && mailDialogHtml"
            class="mail-html meeting-body"
            :sandbox="mailHtmlSandbox"
            referrerpolicy="no-referrer"
            :srcdoc="dialogMail.bodyHtml || dialogMail.body"
          />
          <pre v-else-if="mailHasBody(dialogMail)" class="mail-text meeting-body">{{ dialogMail.body }}</pre>
          <p v-else class="hint meeting-empty-body">目前沒有可顯示的會議內容。</p>
        </div>

        <section class="meeting-attachments">
          <div class="attachment-header">
            <span class="attachment-header-title">
              <span>附件</span>
              <el-tag effect="plain">{{ dialogMailAttachments.length }}</el-tag>
            </span>
          </div>
          <div v-if="isAttachmentListLoading(dialogMail)" class="dialog-loading-block compact">
            <el-skeleton animated :rows="2" />
          </div>
          <p v-else-if="dialogMailAttachments.length === 0" class="hint">這個會議邀請沒有附件。</p>
          <div v-else class="attachment-list">
            <div v-for="attachment in dialogMailAttachments" :key="attachment.attachmentId" class="attachment-row meeting-attachment-row">
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
        </section>
      </main>
    </div>
  </el-dialog>
</template>
