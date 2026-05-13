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
  fetchedMailFolderName,
  flagDisplayLabel,
  flagTagType,
  loadingMails,
  mailCount,
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
  startMailDrag,
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
      @start-mail-drag="startMailDrag"
    />
  </main>
</template>
