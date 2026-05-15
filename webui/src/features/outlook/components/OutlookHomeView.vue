<script setup lang="ts">
import MailListPane from './MailListPane.vue'
import OutlookFolderPane from './OutlookFolderPane.vue'
import type { OutlookDashboardState } from '../composables/useOutlookDashboard'

const props = defineProps<{
  dashboard: OutlookDashboardState
}>()

const {
  categoryTagStyle,
  clearMailDrag,
  deleteMail,
  deleteSelectedMails,
  fetchedMailFolderName,
  flagDisplayLabel,
  flagTagType,
  loadingMails,
  mailCount,
  mailDragPreview,
  mailFetchCountdownText,
  mailFetchStatusText,
  mailListMode,
  mailLookbackHours,
  mails,
  mailStats,
  openMailDialog,
  outlookBusy,
  requestMails,
  selectMail,
  selectedFolderName,
  selectedMailIds,
  showFolderMails,
  showMailFetchWarning,
  splitCategories,
  startMailPointerDrag,
} = props.dashboard
</script>

<template>
  <main class="outlook-layout">
    <OutlookFolderPane :dashboard="dashboard" />

    <MailListPane
      v-model:mail-count="mailCount"
      v-model:mail-lookback-hours="mailLookbackHours"
      :category-tag-style="categoryTagStyle"
      :delete-mail="deleteMail"
      :delete-selected-mails="deleteSelectedMails"
      :fetched-mail-folder-name="fetchedMailFolderName"
      :flag-display-label="flagDisplayLabel"
      :flag-tag-type="flagTagType"
      :loading-mails="loadingMails"
      :mail-fetch-countdown-text="mailFetchCountdownText"
      :mail-fetch-status-text="mailFetchStatusText"
      :mail-list-mode="mailListMode"
      :mail-stats="mailStats"
      :mails="mails"
      :outlook-busy="outlookBusy"
      :selected-folder-name="selectedFolderName"
      :selected-mail-ids="selectedMailIds"
      :show-mail-fetch-warning="showMailFetchWarning"
      :split-categories="splitCategories"
      @clear-mail-drag="clearMailDrag"
      @open-mail-dialog="openMailDialog"
      @request-mails="requestMails"
      @select-mail="selectMail"
      @show-folder-mails="showFolderMails"
      @start-mail-pointer-drag="startMailPointerDrag"
    />

    <teleport to="body">
      <div
        v-if="mailDragPreview.visible"
        class="mail-drag-preview"
        :style="{ left: `${mailDragPreview.x}px`, top: `${mailDragPreview.y}px` }"
      >
        <strong>{{ mailDragPreview.subject }}</strong>
        <span>{{ mailDragPreview.count > 1 ? '多選拖曳' : '移動郵件' }}</span>
      </div>
    </teleport>
  </main>
</template>
