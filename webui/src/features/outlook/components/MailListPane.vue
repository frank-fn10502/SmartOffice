<script setup lang="ts">
import type { CSSProperties } from 'vue'
import { Delete, Document, Rank, View } from '@element-plus/icons-vue'
import type { MailItemDto } from '../models/outlook'
import { formatDateTime } from '../utils/formatters'
import { formatMailSender } from '../utils/mailAddresses'
import { canMoveOutlookItem, outlookItemTypeLabel } from '../utils/outlookItemTypes'

const mailLookbackHours = defineModel<number>('mailLookbackHours', { required: true })
const mailCount = defineModel<number>('mailCount', { required: true })

defineProps<{
  fetchedMailFolderName: string
  mails: MailItemDto[]
  showMailFetchWarning: boolean
  selectedFolderName: string
  mailFetchStatusText: string
  mailStats: { unread: number; flagged: number; categorized: number }
  mailListMode: 'folder' | 'search'
  selectedMailIds: Set<string>
  loadingMails: boolean
  outlookBusy: boolean
  mailFetchCountdownText: string
  categoryTagStyle: (name: string) => CSSProperties
  deleteMail: (mail: MailItemDto) => void
  flagDisplayLabel: (interval: string, request: string) => string
  flagTagType: (interval: string) => string
  splitCategories: (categories: string) => string[]
}>()

defineEmits<{
  clearMailDrag: []
  openMailDialog: [index: number]
  requestMails: []
  selectMail: [index: number, event: MouseEvent]
  showFolderMails: []
  startMailDrag: [mail: MailItemDto, index: number, event: DragEvent]
}>()
</script>

<template>
  <section class="panel outlook-mail-pane">
    <div class="mail-list-toolbar">
      <div class="mail-toolbar-main">
        <div class="panel-title mail-title">
          <el-icon><Document /></el-icon>
          <span>{{ fetchedMailFolderName }}</span>
          <el-tag effect="plain">{{ mails.length }}</el-tag>
          <el-tag v-if="showMailFetchWarning" type="warning" effect="plain">需抓取：{{ selectedFolderName }}</el-tag>
        </div>
        <div class="mail-counts">
          <span>未讀 {{ mailStats.unread }}</span>
          <span>旗標 {{ mailStats.flagged }}</span>
          <span>分類 {{ mailStats.categorized }}</span>
        </div>
        <p class="mail-fetch-status">{{ mailFetchStatusText }}</p>
        <p v-if="showMailFetchWarning" class="mail-fetch-warning">
          目前列表仍是上次抓取的 {{ fetchedMailFolderName }}；已選取 {{ selectedFolderName }}，請按「抓取郵件」更新列表。
        </p>
      </div>
      <div class="mail-toolbar-actions">
        <el-button v-if="mailListMode === 'search'" size="small" @click="$emit('showFolderMails')">回到 folder list</el-button>
        <el-select v-model="mailLookbackHours" class="lookback-select" size="small">
          <el-option label="最近 12 小時" :value="12" />
          <el-option label="最近 24 小時" :value="24" />
          <el-option label="最近 7 天" :value="168" />
          <el-option label="最近 30 天" :value="720" />
          <el-option label="最近 90 天" :value="2160" />
        </el-select>
        <el-select v-model="mailCount" class="count-select" size="small">
          <el-option :value="30" label="30 封" />
          <el-option :value="60" label="60 封" />
          <el-option :value="100" label="100 封" />
        </el-select>
        <el-button type="primary" size="small" :loading="loadingMails" :disabled="outlookBusy && !loadingMails" @click="$emit('requestMails')">
          {{ mailFetchCountdownText ? '立即抓取' : '抓取郵件' }}
        </el-button>
      </div>
    </div>

    <div class="mail-table">
      <p v-if="mails.length === 0 && !loadingMails" class="hint">選取左邊 folder 後抓取郵件。</p>
      <article
        v-for="(mail, index) in mails"
        :key="mail.id || `${mail.receivedTime}-${index}`"
        class="mail-card-row"
        :class="{ selected: selectedMailIds.has(mail.id), unread: !mail.isRead }"
      >
        <div class="mail-row-shell">
          <el-tooltip :content="canMoveOutlookItem(mail) ? '移到刪除的郵件' : '此 Outlook item 不能移動'" placement="top">
            <el-button
              class="mail-delete-button"
              :icon="Delete"
              circle
              size="small"
              type="danger"
              plain
              :disabled="!mail.id?.trim() || outlookBusy || !canMoveOutlookItem(mail)"
              @click.stop="deleteMail(mail)"
            />
          </el-tooltip>
          <el-tooltip :content="canMoveOutlookItem(mail) ? '拖曳移動' : '此 Outlook item 不能移動'" placement="top">
            <button
              class="mail-drag-handle"
              type="button"
              draggable="true"
              :disabled="!mail.id?.trim() || outlookBusy || !canMoveOutlookItem(mail)"
              @click.stop
              @dragstart="$emit('startMailDrag', mail, index, $event)"
              @dragend="$emit('clearMailDrag')"
            >
              <el-icon><Rank /></el-icon>
            </button>
          </el-tooltip>
          <el-tooltip content="開啟郵件" placement="top">
            <el-button
              class="mail-open-button"
              :icon="View"
              circle
              size="small"
              plain
              :disabled="!mail.id?.trim() || outlookBusy"
              @click.stop="$emit('openMailDialog', index)"
            />
          </el-tooltip>
          <button class="mail-row" type="button" @click="$emit('selectMail', index, $event)">
            <span class="mail-row-head">
              <span class="mail-row-main">
                <strong>{{ mail.subject }}</strong>
                <span>{{ formatMailSender(mail) }} · {{ formatDateTime(mail.receivedTime) }}</span>
                <span v-if="mail.attachmentCount > 0" class="mail-row-attachment-summary" :title="mail.attachmentNames">
                  {{ mail.attachmentNames }}
                </span>
              </span>
              <el-tag v-if="mail.attachmentCount > 0" class="mail-attachment-tag" type="info" effect="plain">{{ mail.attachmentCount }} 個附件</el-tag>
            </span>
            <span class="mail-row-tags">
              <el-tag v-if="outlookItemTypeLabel(mail)" type="success" effect="plain">{{ outlookItemTypeLabel(mail) }}</el-tag>
              <el-tag v-if="!mail.isRead" type="warning" effect="plain">未讀</el-tag>
              <el-tag v-if="mail.isMarkedAsTask" :type="flagTagType(mail.flagInterval)" effect="plain">
                {{ flagDisplayLabel(mail.flagInterval, mail.flagRequest) }}<span v-if="mail.taskDueDate"> · {{ formatDateTime(mail.taskDueDate) }}</span>
              </el-tag>
              <el-tag
                v-for="category in splitCategories(mail.categories)"
                :key="category"
                effect="dark"
                :style="categoryTagStyle(category)"
              >
                {{ category }}
              </el-tag>
            </span>
          </button>
        </div>
      </article>
      <div v-if="loadingMails" class="pane-loading">
        <span>Outlook 郵件抓取中...</span>
      </div>
    </div>
  </section>
</template>
