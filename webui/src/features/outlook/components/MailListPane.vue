<script setup lang="ts">
import { onBeforeUnmount, onMounted, ref } from 'vue'
import type { CSSProperties } from 'vue'
import { Document } from '@element-plus/icons-vue'
import type { MailItemDto } from '../models/outlook'
import { formatDateTime } from '../utils/formatters'
import { formatGroupRecipientSummary, groupRecipientsForMail, formatMailSender } from '../utils/mailAddresses'
import { canMoveOutlookItem, outlookItemTypeLabel } from '../utils/outlookItemTypes'

const mailLookbackHours = defineModel<number>('mailLookbackHours', { required: true })
const mailCount = defineModel<number>('mailCount', { required: true })

const props = defineProps<{
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
  deleteSelectedMails: () => void
  flagDisplayLabel: (interval: string, request: string) => string
  flagTagType: (interval: string) => string
  splitCategories: (categories: string) => string[]
}>()

const emit = defineEmits<{
  clearMailDrag: []
  openMailDialog: [index: number]
  requestMails: []
  selectMail: [index: number, event: MouseEvent]
  showFolderMails: []
  startMailPointerDrag: [mail: MailItemDto, index: number, event: PointerEvent]
}>()

const contextMenu = ref({
  visible: false,
  x: 0,
  y: 0,
  mail: null as MailItemDto | null,
  index: -1,
})

function closeMailContextMenu() {
  contextMenu.value.visible = false
}

function openMailContextMenu(mail: MailItemDto, index: number, event: MouseEvent) {
  event.preventDefault()
  if (!mail.id || !props.selectedMailIds.has(mail.id)) emit('selectMail', index, event)
  contextMenu.value = {
    visible: true,
    x: event.clientX,
    y: event.clientY,
    mail,
    index,
  }
}

function contextDeleteLabel() {
  return '刪除郵件'
}

function contextDeleteDisabled() {
  const mail = contextMenu.value.mail
  return !mail?.id?.trim() || props.outlookBusy || !canMoveOutlookItem(mail)
}

function deleteFromContextMenu() {
  const mail = contextMenu.value.mail
  if (!mail) return
  closeMailContextMenu()
  if (mail.id && props.selectedMailIds.size > 1 && props.selectedMailIds.has(mail.id)) {
    props.deleteSelectedMails()
    return
  }
  props.deleteMail(mail)
}

onMounted(() => {
  window.addEventListener('click', closeMailContextMenu)
  window.addEventListener('blur', closeMailContextMenu)
})

onBeforeUnmount(() => {
  window.removeEventListener('click', closeMailContextMenu)
  window.removeEventListener('blur', closeMailContextMenu)
})
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
        :class="{ selected: selectedMailIds.has(mail.id), unread: !mail.isRead, 'has-group-recipient': groupRecipientsForMail(mail).length > 0 }"
      >
        <div class="mail-row-shell">
          <button
            class="mail-row"
            type="button"
            :title="canMoveOutlookItem(mail) ? '按住拖曳到左側 folder 可移動；雙擊開啟郵件。' : '此 Outlook item 不能移動；雙擊可開啟郵件。'"
            @pointerdown="$emit('startMailPointerDrag', mail, index, $event)"
            @click="$emit('selectMail', index, $event)"
            @dblclick="$emit('openMailDialog', index)"
            @contextmenu="openMailContextMenu(mail, index, $event)"
          >
            <span class="mail-row-head">
              <span class="mail-row-main">
                <strong>{{ mail.subject }}</strong>
                <span>{{ formatMailSender(mail) }} · {{ formatDateTime(mail.receivedTime) }}</span>
                <span v-if="mail.attachmentCount > 0" class="mail-row-attachment-summary">
                  {{ mail.attachmentNames }}
                </span>
                <span v-if="groupRecipientsForMail(mail).length > 0" class="mail-row-group-summary">
                  Group 收件者：{{ formatGroupRecipientSummary(mail) }}
                </span>
              </span>
              <el-tag v-if="mail.attachmentCount > 0" class="mail-attachment-tag" type="info" effect="plain">{{ mail.attachmentCount }} 個附件</el-tag>
            </span>
            <span class="mail-row-tags">
              <el-tag v-if="groupRecipientsForMail(mail).length > 0" type="info" effect="plain">
                Group {{ groupRecipientsForMail(mail).length }}
              </el-tag>
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

    <teleport to="body">
      <div
        v-if="contextMenu.visible"
        class="mail-context-menu"
        :style="{ left: `${contextMenu.x}px`, top: `${contextMenu.y}px` }"
        @click.stop
      >
        <button type="button" class="danger" :disabled="contextDeleteDisabled()" @click="deleteFromContextMenu">
          {{ contextDeleteLabel() }}
        </button>
      </div>
    </teleport>
  </section>
</template>
