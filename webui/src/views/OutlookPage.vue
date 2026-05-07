<script setup lang="ts">
import {
  ArrowLeft,
  ArrowRight,
  Calendar,
  ChatDotRound,
  Delete,
  Document,
  Folder,
  Rank,
  Refresh,
  Search,
  View,
} from '@element-plus/icons-vue'
import CategoryManagerDialog from '../components/CategoryManagerDialog.vue'
import FlagEditorDialog from '../components/FlagEditorDialog.vue'
import FolderNode from '../components/FolderNode.vue'
import MailPropertyPane from '../components/MailPropertyPane.vue'
import type { OutlookDashboardState } from '../composables/useOutlookDashboard'
import type { AppView } from '../models/outlook'
import { formatDateTime, formatTime } from '../utils/formatters'
import { formatMailSender, formatRecipient, formatRecipients, shouldShowRecipientSmtpAddress } from '../utils/mailAddresses'

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
  activeView,
  addCategoryToMasterList,
  addinLogs,
  addinStatus,
  applyMailProperties,
  attachmentExportRootDraft,
  attachmentExportSettings,
  calendarEvents,
  calendarMonthLabel,
  calendarWeekdays,
  calendarWeeks,
  cancelCreateFolder,
  categories,
  categoryManagerVisible,
  categoryColorOptions,
  categoryColorStyle,
  categoryTagStyle,
  categoryCreateColor,
  categoryCreateDraft,
  changeCalendarMonth,
  chatMessages,
  chatPanelRef,
  chatText,
  closeMailDialog,
  clearMailDrag,
  contextFolderName,
  createFolder,
  createFolderFromContext,
  creatingFolderName,
  creatingFolderParentPath,
  deleteFolderFromContext,
  deleteMail,
  dragOverFolderPath,
  draggedMailId,
  expandedFolders,
  exportMailAttachment,
  fetchMailsFromContext,
  fetchedMailFolderName,
  flagDisplayLabel,
  flagIntervalOptions,
  flagTagType,
  folderContextMenu,
  folderStores,
  isAttachmentExporting,
  isAttachmentListLoading,
  isMailBodyLoading,
  loadingCalendar,
  loadingCategories,
  loadingFolders,
  loadingMailSearch,
  loadingMails,
  loadingSignalRPing,
  flagEditorVisible,
  mailCount,
  mailHtmlSandbox,
  mailListMode,
  mailListNeedsFetch,
  mailFetchCountdownText,
  showMailFetchWarning,
  mailFetchStatusText,
  mailHasBody,
  mailPropertiesDraft,
  mailPropertiesChanged,
  mailLookbackHours,
  mailSearchDraft,
  mailSearchProgressText,
  mailSearchSummaryItems,
  mailStats,
  searchResultGroups,
  searchResultRows,
  searchResultViewMode,
  dialogMail,
  dialogMailAttachments,
  dialogMailFolderName,
  dialogMailHasIdentity,
  dialogLoading,
  mailDialogHtml,
  mailDialogVisible,
  masterCategoryListExpanded,
  mails,
  moveDraggedMail,
  navOptions,
  openExportedAttachment,
  openFolderContextMenu,
  operationLoading,
  openCategoryManager,
  outlookBusy,
  outlookBusyText,
  openMailDialog,
  refreshAdminData,
  requestCalendar,
  requestCategories,
  requestFolders,
  requestSignalRPing,
  requestMails,
  requestMailSearch,
  resetMailPropertiesDraft,
  resetAttachmentExportRoot,
  saveAttachmentExportSettings,
  savingAttachmentExportSettings,
  selectedFolderName,
  selectedFolderPath,
  selectedCalendarEvent,
  selectedMailIndex,
  selectedMailIds,
  selectFolder,
  selectCalendarEvent,
  selectMail,
  sendChat,
  showFolderMails,
  goToCurrentCalendarMonth,
  addMailCategoryDraft,
  removeMailCategoryDraft,
  setDragOverFolder,
  setMailFlagDraft,
  signalRState,
  splitCategories,
  startMailDrag,
  switchView,
  toggleFolder,
  toggleMasterCategoryList,
  toggleSearchResultFolder,
  toggleSearchResultStore,
  updateCategoryColor,
  visibleFolders,
  visibleMasterCategories,
  hiddenMasterCategoryCount,
} = props.dashboard
</script>

<template>
  <div class="outlook-page">
      <div class="feature-toolbar">
        <div class="feature-title">
          <el-icon><Document /></el-icon>
          <span>Outlook</span>
          <el-tag :type="signalRState === 'connected' ? 'success' : 'danger'" effect="plain">
            {{ signalRState }}
          </el-tag>
        </div>

        <el-segmented
          :model-value="activeView"
          :options="navOptions"
          @update:model-value="(value: string | number | boolean) => switchView(value as AppView)"
        />
      </div>

      <main v-if="activeView === 'outlook'" class="outlook-layout">
        <section class="panel outlook-folder-pane">
          <div class="panel-header">
            <div class="panel-title">
              <el-icon><Folder /></el-icon>
              <span>Folders</span>
            </div>
            <el-button :icon="Refresh" circle :loading="loadingFolders" :disabled="outlookBusy && !loadingFolders" @click="requestFolders" />
          </div>

          <div class="folder-list outlook-folder-list">
            <p v-if="visibleFolders.length === 0 && !loadingFolders" class="hint">Waiting for folders...</p>
            <FolderNode
              v-for="folder in visibleFolders"
              :key="folder.folderPath"
              :folder="folder"
              :store="folderStores.find((store) => store.storeId === folder.storeId)"
              :level="0"
              :expanded-folders="expandedFolders"
              :selected-folder-path="selectedFolderPath"
              :creating-folder-parent-path="creatingFolderParentPath"
              :creating-folder-name="creatingFolderName"
              :folder-busy="outlookBusy"
              :can-drop-mail="Boolean(draggedMailId) && !outlookBusy"
              :active-drop-folder-path="dragOverFolderPath"
              @toggle="toggleFolder"
              @select="selectFolder"
              @context="openFolderContextMenu"
              @update:creating-folder-name="creatingFolderName = $event"
              @create="createFolder($event.parentPath, $event.name)"
              @cancel-create="cancelCreateFolder"
              @drag-mail-over="setDragOverFolder"
              @drop-mail="moveDraggedMail"
            />
            <div v-if="loadingFolders" class="pane-loading">
              <span>Outlook folder 同步中...</span>
            </div>
          </div>
        </section>

        <section class="panel outlook-mail-pane">
          <div class="panel-header mail-list-header">
            <div class="mail-header-main">
              <div class="panel-title mail-title">
                <el-icon><Document /></el-icon>
                <span>{{ fetchedMailFolderName }}</span>
                <el-tag effect="plain">{{ mails.length }}</el-tag>
                <el-tag v-if="showMailFetchWarning" type="warning" effect="plain">需抓取：{{ selectedFolderName }}</el-tag>
              </div>
              <p class="mail-fetch-status">{{ mailFetchStatusText }}</p>
              <p v-if="showMailFetchWarning" class="mail-fetch-warning">
                目前列表仍是上次抓取的 {{ fetchedMailFolderName }}；已選取 {{ selectedFolderName }}，請按「抓取郵件」更新列表。
              </p>
            </div>

            <div class="mail-header-actions">
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
              <el-button type="primary" size="small" :loading="loadingMails" :disabled="outlookBusy && !loadingMails" @click="requestMails">
                {{ mailFetchCountdownText ? '立即抓取' : '抓取郵件' }}
              </el-button>
            </div>
          </div>

          <div class="mail-fetch-bar">
            <div class="mail-counts">
              <span>未讀 {{ mailStats.unread }}</span>
              <span>旗標 {{ mailStats.flagged }}</span>
              <span>分類 {{ mailStats.categorized }}</span>
            </div>
            <el-button v-if="mailListMode === 'search'" size="small" @click="showFolderMails">回到 folder list</el-button>
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
                <el-tooltip content="刪除郵件" placement="top">
                  <el-button
                    class="mail-delete-button"
                    :icon="Delete"
                    circle
                    size="small"
                    type="danger"
                    plain
                    :disabled="!mail.id?.trim() || outlookBusy"
                    @click.stop="deleteMail(mail)"
                  />
                </el-tooltip>
                <el-tooltip content="拖曳移動郵件" placement="top">
                  <button
                    class="mail-drag-handle"
                    type="button"
                    draggable="true"
                    :disabled="!mail.id?.trim() || outlookBusy"
                    @click.stop
                    @dragstart="startMailDrag(mail, index, $event)"
                    @dragend="clearMailDrag"
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
                    @click.stop="openMailDialog(index)"
                  />
                </el-tooltip>
                <button
                  class="mail-row"
                  type="button"
                  @click="selectMail(index, $event)"
                >
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
      </main>

      <main v-else-if="activeView === 'search'" class="search-layout">
        <section class="panel outlook-folder-pane">
          <div class="panel-header">
            <div class="panel-title">
              <el-icon><Folder /></el-icon>
              <span>Folders</span>
            </div>
            <el-button :icon="Refresh" circle :loading="loadingFolders" :disabled="outlookBusy && !loadingFolders" @click="requestFolders" />
          </div>

          <div class="folder-list outlook-folder-list">
            <p v-if="visibleFolders.length === 0 && !loadingFolders" class="hint">Waiting for folders...</p>
            <FolderNode
              v-for="folder in visibleFolders"
              :key="folder.folderPath"
              :folder="folder"
              :store="folderStores.find((store) => store.storeId === folder.storeId)"
              :level="0"
              :expanded-folders="expandedFolders"
              :selected-folder-path="selectedFolderPath"
              :creating-folder-parent-path="creatingFolderParentPath"
              :creating-folder-name="creatingFolderName"
              :folder-busy="outlookBusy"
              :can-drop-mail="Boolean(draggedMailId) && !outlookBusy"
              :active-drop-folder-path="dragOverFolderPath"
              @toggle="toggleFolder"
              @select="selectFolder"
              @context="openFolderContextMenu"
              @update:creating-folder-name="creatingFolderName = $event"
              @create="createFolder($event.parentPath, $event.name)"
              @cancel-create="cancelCreateFolder"
              @drag-mail-over="setDragOverFolder"
              @drop-mail="moveDraggedMail"
            />
            <div v-if="loadingFolders" class="pane-loading">
              <span>Outlook folder 同步中...</span>
            </div>
          </div>
        </section>

        <section class="panel search-page-panel">
          <div class="panel-header">
            <div class="panel-title">
              <el-icon><Search /></el-icon>
              <span>搜尋郵件</span>
              <el-tag effect="plain">{{ mails.length }}</el-tag>
            </div>
            <el-button type="primary" :loading="loadingMailSearch" :disabled="loadingMailSearch" @click="requestMailSearch">
              搜尋
            </el-button>
          </div>

          <div class="mail-search-panel standalone-search-panel">
            <div class="mail-search-flow">
              使用 Outlook 內建搜尋尋找符合文字與篩選條件的郵件；不做錯字模糊比對，也不限制結果筆數。
            </div>
            <div class="mail-search-row">
              <el-input
                v-model="mailSearchDraft.keyword"
                clearable
                :prefix-icon="Search"
                placeholder="片段關鍵字，例如：客戶xxxx"
                @keydown.enter.prevent="requestMailSearch"
              />
              <el-select v-model="mailSearchDraft.scopeMode" class="scope-select">
                <el-option label="目前資料夾" value="selected_folder" />
                <el-option label="目前信箱" value="selected_store" />
                <el-option label="全部信箱" value="global" />
              </el-select>
            </div>
            <div class="mail-search-row search-options-row">
              <span class="search-options-label">文字範圍</span>
              <el-checkbox-group v-model="mailSearchDraft.textFields">
                <el-checkbox label="subject">標題</el-checkbox>
                <el-checkbox label="sender">寄件者</el-checkbox>
                <el-checkbox label="body">內容</el-checkbox>
              </el-checkbox-group>
            </div>
            <div class="mail-search-row search-options-row">
              <span class="search-options-label">篩選條件</span>
              <el-select
                v-model="mailSearchDraft.categoryNames"
                class="category-filter-select"
                multiple
                collapse-tags
                collapse-tags-tooltip
                clearable
                placeholder="分類"
              >
                <el-option
                  v-for="category in categories"
                  :key="category.name"
                  :label="category.name"
                  :value="category.name"
                />
              </el-select>
              <el-select v-model="mailSearchDraft.hasAttachments" class="filter-select" clearable placeholder="附件">
                <el-option label="包含附件" :value="true" />
                <el-option label="不含附件" :value="false" />
              </el-select>
              <el-select v-model="mailSearchDraft.flagState" class="filter-select">
                <el-option label="旗標不限" value="any" />
                <el-option label="有旗標" value="flagged" />
                <el-option label="無旗標" value="unflagged" />
              </el-select>
              <el-select v-model="mailSearchDraft.readState" class="filter-select">
                <el-option label="已讀不限" value="any" />
                <el-option label="未讀" value="unread" />
                <el-option label="已讀" value="read" />
              </el-select>
            </div>
            <div class="mail-search-row search-options-row">
              <el-date-picker
                v-model="mailSearchDraft.receivedFrom"
                type="datetime"
                value-format="YYYY-MM-DDTHH:mm:ss"
                placeholder="收到時間起"
              />
              <el-date-picker
                v-model="mailSearchDraft.receivedTo"
                type="datetime"
                value-format="YYYY-MM-DDTHH:mm:ss"
                placeholder="收到時間迄"
              />
            </div>
          </div>

          <div class="search-result-criteria" :class="{ empty: mailSearchSummaryItems.length === 0 }">
            <div class="search-result-summary" :class="{ empty: mailSearchSummaryItems.length === 0 }">
              <template v-if="mailSearchSummaryItems.length > 0">
                <span
                  v-for="item in mailSearchSummaryItems"
                  :key="`${item.label}-${item.value}`"
                  class="search-summary-chip"
                  :class="item.tone"
                >
                  <span>{{ item.label }}</span>
                  <strong>{{ item.value }}</strong>
                </span>
              </template>
              <span v-else>尚未送出搜尋條件</span>
            </div>
          </div>

          <div class="search-result-toolbar">
            <div class="search-result-total">
              搜尋結果 <strong>{{ mails.length }}</strong> 封
            </div>
            <el-segmented
              v-model="searchResultViewMode"
              :options="[
                { label: 'Tree', value: 'tree' },
                { label: 'Flat', value: 'flat' },
              ]"
            />
          </div>
          <div v-if="loadingMailSearch" class="search-result-loading" role="status">
            <span>{{ mailSearchProgressText || 'Outlook 郵件搜尋中...' }}</span>
          </div>

          <div class="mail-table search-result-table">
            <p v-if="mails.length === 0 && !loadingMailSearch" class="hint">尚未取得搜尋結果。</p>
            <article
              v-for="{ mail, index, sourceLabel } in searchResultViewMode === 'flat' ? searchResultRows : []"
              :key="mail.id || `${mail.receivedTime}-${index}`"
              class="mail-card-row"
              :class="{ selected: selectedMailIds.has(mail.id), unread: !mail.isRead }"
            >
              <div class="mail-row-shell">
                <el-tooltip content="拖曳移動郵件" placement="top">
                  <button
                    class="mail-drag-handle"
                    type="button"
                    draggable="true"
                    :disabled="!mail.id?.trim() || outlookBusy"
                    @click.stop
                    @dragstart="startMailDrag(mail, index, $event)"
                    @dragend="clearMailDrag"
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
                    @click.stop="openMailDialog(index)"
                  />
                </el-tooltip>
                <button
                  class="mail-row"
                  type="button"
                  @click="selectMail(index, $event)"
                  @dblclick="openMailDialog(index)"
                >
                  <span class="mail-row-head">
                    <span class="mail-row-main">
                      <strong>{{ mail.subject }}</strong>
                      <span>{{ formatMailSender(mail) }} · {{ formatDateTime(mail.receivedTime) }}</span>
                      <span class="mail-source-label">{{ sourceLabel }}</span>
                    </span>
                    <el-tag v-if="mail.attachmentCount > 0" class="mail-attachment-tag" type="info" effect="plain">{{ mail.attachmentCount }} 個附件</el-tag>
                  </span>
                  <span class="mail-row-tags">
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
            <div v-for="store in searchResultViewMode === 'tree' ? searchResultGroups : []" :key="store.key" class="search-result-store">
              <button class="search-result-tree-node store" type="button" @click="toggleSearchResultStore(store.key)">
                <el-icon class="search-result-disclosure" :class="{ collapsed: store.collapsed }"><ArrowRight /></el-icon>
                <el-icon><Folder /></el-icon>
                <span class="search-result-node-label">{{ store.label }}</span>
                <span class="search-result-count">{{ store.count }}</span>
              </button>
              <div v-if="!store.collapsed" class="search-result-store-children">
                <div v-for="folderGroup in store.folders" :key="folderGroup.key" class="search-result-folder">
                  <button class="search-result-tree-node folder" type="button" @click="toggleSearchResultFolder(folderGroup.key)">
                    <el-icon class="search-result-disclosure" :class="{ collapsed: folderGroup.collapsed }"><ArrowRight /></el-icon>
                    <el-icon><Folder /></el-icon>
                    <span class="search-result-node-label">{{ folderGroup.label }}</span>
                    <span v-if="folderGroup.path" class="search-result-node-path">{{ folderGroup.path }}</span>
                    <span class="search-result-count">{{ folderGroup.count }}</span>
                  </button>
                  <div v-if="!folderGroup.collapsed" class="search-result-folder-children">
                    <article
                      v-for="{ mail, index } in folderGroup.rows"
                      :key="mail.id || `${mail.receivedTime}-${index}`"
                      class="mail-card-row search-result-tree-row"
                      :class="{ selected: selectedMailIds.has(mail.id), unread: !mail.isRead }"
                    >
                      <div class="mail-row-shell">
                        <el-tooltip content="拖曳移動郵件" placement="top">
                          <button
                            class="mail-drag-handle"
                            type="button"
                            draggable="true"
                            :disabled="!mail.id?.trim() || outlookBusy"
                            @click.stop
                            @dragstart="startMailDrag(mail, index, $event)"
                            @dragend="clearMailDrag"
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
                            @click.stop="openMailDialog(index)"
                          />
                        </el-tooltip>
                        <button
                          class="mail-row"
                          type="button"
                          @click="selectMail(index, $event)"
                          @dblclick="openMailDialog(index)"
                        >
                          <span class="mail-row-head">
                            <span class="mail-row-main">
                              <strong>{{ mail.subject }}</strong>
                              <span>{{ formatMailSender(mail) }} · {{ formatDateTime(mail.receivedTime) }}</span>
                            </span>
                            <el-tag v-if="mail.attachmentCount > 0" class="mail-attachment-tag" type="info" effect="plain">{{ mail.attachmentCount }} 個附件</el-tag>
                          </span>
                          <span class="mail-row-tags">
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
                  </div>
                </div>
              </div>
            </div>
          </div>
        </section>
      </main>

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

      <main v-if="activeView === 'chat'" class="chat-layout">
        <section class="panel chat-page-panel">
          <div class="panel-header">
            <div class="panel-title">
              <el-icon><ChatDotRound /></el-icon>
              <span>Chat</span>
            </div>
          </div>

          <div ref="chatPanelRef" class="chat-messages">
            <div
              v-for="(message, index) in chatMessages"
              :key="message.id ?? `${message.timestamp}-${index}`"
              class="chat-message"
              :class="{ web: message.source === 'web' }"
            >
              <span class="chat-meta">[{{ message.source }}] {{ formatTime(message.timestamp) }}</span>
              <span class="chat-bubble">{{ message.text }}</span>
            </div>
          </div>

          <div class="chat-input">
            <el-input v-model="chatText" placeholder="Send message..." @keydown.enter="sendChat" />
            <el-button type="primary" @click="sendChat">Send</el-button>
          </div>
        </section>
      </main>

      <main v-if="activeView === 'calendar'" class="calendar-layout">
        <section class="panel">
          <div class="panel-header">
            <div class="panel-title">
              <el-icon><Calendar /></el-icon>
              <span>月曆</span>
              <el-tag effect="plain">{{ calendarEvents.length }}</el-tag>
            </div>
            <div class="calendar-actions">
              <el-button :icon="ArrowLeft" :disabled="outlookBusy" @click="changeCalendarMonth(-1)" />
              <strong>{{ calendarMonthLabel }}</strong>
              <el-button :disabled="outlookBusy" @click="goToCurrentCalendarMonth">本月</el-button>
              <el-button :icon="ArrowRight" :disabled="outlookBusy" @click="changeCalendarMonth(1)" />
              <el-button :icon="Refresh" :loading="loadingCalendar" :disabled="outlookBusy && !loadingCalendar" @click="requestCalendar">
                同步整月
              </el-button>
            </div>
          </div>

          <div class="calendar-page">
            <div class="calendar-grid">
              <div v-for="day in calendarWeekdays" :key="day" class="calendar-weekday">{{ day }}</div>
              <div v-for="week in calendarWeeks" :key="week.key" class="calendar-week-row">
                <div class="calendar-week-days">
                  <div
                    v-for="day in week.days"
                    :key="day.key"
                    class="calendar-day"
                    :class="{ muted: !day.inMonth, today: day.isToday }"
                  >
                    <div class="calendar-day-number">{{ day.dayNumber }}</div>
                  </div>
                </div>
                <div class="calendar-week-events">
                  <button
                    v-for="segment in week.segments"
                    :key="`${segment.event.id || segment.event.start}-${segment.startColumn}`"
                    class="calendar-event"
                    :class="{ continued: segment.isMultiDay, 'continues-before': !segment.isStart, 'continues-after': !segment.isEnd }"
                    type="button"
                    :style="{ gridColumn: `${segment.startColumn} / span ${segment.span}` }"
                    @click="selectCalendarEvent(segment.event)"
                  >
                    <span>{{ segment.isMultiDay ? `${formatDateTime(segment.event.start)} - ${formatDateTime(segment.event.end)}` : formatTime(segment.event.start) }}</span>
                    <strong>{{ segment.event.subject }}</strong>
                  </button>
                </div>
              </div>
            </div>

            <aside class="calendar-detail">
              <template v-if="selectedCalendarEvent">
                <div class="calendar-detail-title">{{ selectedCalendarEvent.subject }}</div>
                <div class="rule-detail">
                  <span>{{ formatDateTime(selectedCalendarEvent.start) }} - {{ formatDateTime(selectedCalendarEvent.end) }}</span>
                  <span>地點：{{ selectedCalendarEvent.location || '-' }}</span>
                  <span>召集人：{{ formatRecipient(selectedCalendarEvent.organizer, '-') }}</span>
                  <span>出席者：{{ formatRecipients(selectedCalendarEvent.requiredAttendees) || '-' }}</span>
                </div>
                <div class="marker-tags">
                  <el-tag effect="plain">{{ selectedCalendarEvent.busyStatus || 'unknown' }}</el-tag>
                  <el-tag v-if="selectedCalendarEvent.isRecurring" type="warning" effect="plain">週期性</el-tag>
                </div>
              </template>
              <div v-else class="empty-inspector">
                點選月曆中的項目查看詳細資訊。
              </div>
            </aside>
          </div>
        </section>
      </main>

      <div
        v-if="folderContextMenu.visible"
        class="folder-context-menu"
        :style="{ left: `${folderContextMenu.x}px`, top: `${folderContextMenu.y}px` }"
        @click.stop
      >
        <div class="context-menu-title">{{ contextFolderName }}</div>
        <button type="button" :disabled="outlookBusy" @click="fetchMailsFromContext">抓取郵件</button>
        <button type="button" :disabled="outlookBusy" @click="createFolderFromContext">新增子資料夾</button>
        <button class="danger" type="button" :disabled="outlookBusy" @click="deleteFolderFromContext">刪除此資料夾</button>
      </div>

      <CategoryManagerDialog
        v-model="categoryManagerVisible"
        v-model:category-create-color="categoryCreateColor"
        v-model:category-create-draft="categoryCreateDraft"
        :categories="categories"
        :category-color-options="categoryColorOptions"
        :category-color-style="categoryColorStyle"
        :hidden-master-category-count="hiddenMasterCategoryCount"
        :loading-categories="loadingCategories"
        :master-category-list-expanded="masterCategoryListExpanded"
        :operation-loading="operationLoading"
        :outlook-busy="outlookBusy"
        :visible-master-categories="visibleMasterCategories"
        @add-category="addCategoryToMasterList"
        @request-categories="requestCategories"
        @toggle-master-category-list="toggleMasterCategoryList"
        @update-category-color="updateCategoryColor"
      />

      <FlagEditorDialog
        v-model="flagEditorVisible"
        v-model:draft="mailPropertiesDraft"
        :outlook-busy="outlookBusy"
      />
  </div>
</template>
