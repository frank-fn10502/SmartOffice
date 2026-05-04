<script setup lang="ts">
import {
  ArrowLeft,
  ArrowRight,
  Calendar,
  ChatDotRound,
  Connection,
  Document,
  Folder,
  Monitor,
  PriceTag,
  Refresh,
  Search,
} from '@element-plus/icons-vue'
import FolderNode from './components/FolderNode.vue'
import { useOutlookDashboard } from './composables/useOutlookDashboard'
import type { AppView } from './models/outlook'
import { formatDateTime, formatTime } from './utils/formatters'

function formatAttachmentSize(size: number) {
  if (size >= 1024 * 1024) return `${(size / 1024 / 1024).toFixed(1)} MB`
  if (size >= 1024) return `${Math.round(size / 1024)} KB`
  return `${size} B`
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
  categoryColorOptions,
  categoryColorStyle,
  categoryTagStyle,
  categoryCreateColor,
  categoryCreateDraft,
  changeCalendarMonth,
  chatMessages,
  chatPanelRef,
  chatText,
  clearMailDrag,
  contextFolderName,
  createFolder,
  createFolderFromContext,
  creatingFolderName,
  creatingFolderParentPath,
  deleteFolderFromContext,
  dragOverFolderPath,
  draggedMailId,
  expandedFolders,
  exportMailAttachment,
  fetchMailsFromContext,
  fetchedMailFolderName,
  flagIntervalLabel,
  flagIntervalOptions,
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
  mailCount,
  mailHtmlSandbox,
  mailListMode,
  mailListNeedsFetch,
  mailHasBody,
  mailPropertiesDraft,
  mailRange,
  mailSearchDraft,
  mailStats,
  mails,
  moveDraggedMail,
  openExportedAttachment,
  openFolderContextMenu,
  operationLoading,
  outlookBusy,
  outlookBusyText,
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
  selectedMail,
  selectedMailAttachments,
  selectedMailCategories,
  selectedMailFolderName,
  selectedMailHasIdentity,
  selectedMailHtml,
  selectedMailIndex,
  selectedMailIsOpen,
  selectFolder,
  selectCalendarEvent,
  selectMail,
  sendChat,
  showFolderMails,
  goToCurrentCalendarMonth,
  setDragOverFolder,
  signalRState,
  splitCategories,
  startMailDrag,
  switchView,
  toggleFolder,
  updateCategoryColor,
  visibleFolders,
} = useOutlookDashboard()
</script>

<template>
  <el-config-provider size="default">
    <div class="app-shell">
      <header class="topbar">
        <div class="brand">
          <el-icon><Monitor /></el-icon>
          <span>SmartOffice Dashboard</span>
          <el-tag :type="signalRState === 'connected' ? 'success' : 'danger'" effect="dark">
            {{ signalRState }}
          </el-tag>
        </div>

        <nav class="nav-actions">
          <el-segmented
            :model-value="activeView"
            :options="[
              { label: 'Outlook', value: 'outlook' },
              { label: 'Chat', value: 'chat' },
              { label: 'Calendar', value: 'calendar' },
              { label: 'Admin', value: 'admin' },
              { label: 'Swagger', value: 'swagger' },
            ]"
            @update:model-value="(value: string | number | boolean) => switchView(value as AppView)"
          />
        </nav>
      </header>

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
          <div class="panel-header">
            <div class="panel-title">
              <el-icon><Document /></el-icon>
              <span>{{ fetchedMailFolderName }}</span>
              <el-tag effect="plain">{{ mails.length }}</el-tag>
              <el-tag v-if="mailListNeedsFetch" type="warning" effect="plain">需抓取：{{ selectedFolderName }}</el-tag>
            </div>

            <el-button type="primary" :loading="loadingMails" :disabled="outlookBusy && !loadingMails" @click="requestMails">
              抓取郵件
            </el-button>
          </div>

          <div class="mail-fetch-bar">
            <el-select v-model="mailRange" class="range-select">
              <el-option label="今天" value="1d" />
              <el-option label="最近 7 天" value="1w" />
              <el-option label="最近 30 天" value="1m" />
            </el-select>
            <el-select v-model="mailCount" class="count-select">
              <el-option :value="10" label="10" />
              <el-option :value="20" label="20" />
              <el-option :value="30" label="30" />
              <el-option :value="100" label="100" />
            </el-select>
            <div class="mail-counts">
              <span>未讀 {{ mailStats.unread }}</span>
              <span>旗標 {{ mailStats.flagged }}</span>
              <span>分類 {{ mailStats.categorized }}</span>
            </div>
            <el-button v-if="mailListMode === 'search'" size="small" @click="showFolderMails">回到 folder list</el-button>
          </div>

          <div class="mail-search-panel">
            <div class="mail-search-row">
              <el-input
                v-model="mailSearchDraft.keyword"
                clearable
                :prefix-icon="Search"
                placeholder="片段關鍵字，例如：客戶xxxx"
                @keydown.enter.prevent="requestMailSearch"
              />
              <el-select v-model="mailSearchDraft.matchMode" class="match-select">
                <el-option label="片段" value="contains" />
                <el-option label="完全相同" value="exact" />
                <el-option label="Regex 後篩" value="regex" />
              </el-select>
              <el-select v-model="mailSearchDraft.scopeMode" class="scope-select">
                <el-option label="目前 folder" value="selected_folder" />
                <el-option label="目前 store" value="selected_store" />
                <el-option label="全域分批" value="global" />
              </el-select>
              <el-select v-model="mailSearchDraft.maxCount" class="count-select">
                <el-option :value="25" label="25" />
                <el-option :value="50" label="50" />
                <el-option :value="100" label="100" />
                <el-option :value="200" label="200" />
              </el-select>
              <el-button type="primary" :loading="loadingMailSearch" :disabled="loadingMailSearch" @click="requestMailSearch">
                搜尋
              </el-button>
            </div>
            <div class="mail-search-row search-options-row">
              <el-date-picker
                v-model="mailSearchDraft.receivedFrom"
                type="datetime"
                value-format="YYYY-MM-DDTHH:mm:ss"
                placeholder="起始時間"
              />
              <el-date-picker
                v-model="mailSearchDraft.receivedTo"
                type="datetime"
                value-format="YYYY-MM-DDTHH:mm:ss"
                placeholder="結束時間"
              />
              <el-date-picker
                v-model="mailSearchDraft.exactReceivedTime"
                type="datetime"
                value-format="YYYY-MM-DDTHH:mm:ss"
                placeholder="單一時間"
              />
              <el-input-number
                v-model="mailSearchDraft.exactReceivedToleranceSeconds"
                :min="0"
                :max="3600"
                :step="30"
                controls-position="right"
                class="tolerance-input"
              />
              <el-checkbox v-model="mailSearchDraft.includeSubFolders">含子 folder</el-checkbox>
            </div>
            <div class="mail-search-row search-options-row">
              <el-checkbox-group v-model="mailSearchDraft.fields">
                <el-checkbox label="subject">Subject</el-checkbox>
                <el-checkbox label="sender">寄件者</el-checkbox>
                <el-checkbox label="categories">分類</el-checkbox>
                <el-checkbox label="body">Body</el-checkbox>
              </el-checkbox-group>
              <span class="search-note">Regex 只適合小範圍後篩；大型搜尋請優先使用時間、folder 與 store 條件。</span>
            </div>
          </div>
          <p v-if="mailListNeedsFetch" class="hint">
            目前列表仍是上次抓取的 {{ fetchedMailFolderName }}；已選取 {{ selectedFolderName }}，請按「抓取郵件」更新列表。
          </p>

          <div class="mail-table">
            <p v-if="mails.length === 0 && !loadingMails" class="hint">選取左邊 folder 後抓取郵件。</p>
            <article
              v-for="(mail, index) in mails"
              :key="mail.id || `${mail.receivedTime}-${index}`"
              class="mail-card-row"
              :class="{ selected: selectedMailIndex === index, unread: !mail.isRead }"
            >
              <button
                class="mail-row"
                type="button"
                draggable="true"
                @click="selectMail(index)"
                @dragstart="startMailDrag(mail, index, $event)"
                @dragend="clearMailDrag"
              >
                <span class="mail-row-main">
                  <strong>{{ mail.subject }}</strong>
                  <span>{{ mail.senderName }} · {{ formatDateTime(mail.receivedTime) }}</span>
                </span>
                <span class="mail-row-tags">
                  <el-tag v-if="!mail.isRead" type="warning" effect="plain">未讀</el-tag>
                  <el-tag v-if="mail.isMarkedAsTask" type="danger" effect="plain">
                    {{ flagIntervalLabel(mail.flagInterval) }}
                  </el-tag>
                  <el-tag v-if="mail.taskDueDate" type="info" effect="plain">到期 {{ formatDateTime(mail.taskDueDate) }}</el-tag>
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

              <div v-if="selectedMailIndex === index && selectedMailIsOpen" class="mail-inline-detail">
                <div v-if="isMailBodyLoading(mail)" class="pane-loading">
                  <span>郵件內容載入中...</span>
                </div>
                <el-button v-else-if="mailHasBody(mail)" size="small" @click="selectedMailHtml = !selectedMailHtml">
                  {{ selectedMailHtml ? '切到文字' : '切到 HTML' }}
                </el-button>
                <iframe
                  v-if="mailHasBody(mail) && selectedMailHtml"
                  class="mail-html"
                  :sandbox="mailHtmlSandbox"
                  referrerpolicy="no-referrer"
                  :srcdoc="mail.bodyHtml || mail.body"
                />
                <pre v-else-if="mailHasBody(mail)" class="mail-text">{{ mail.body }}</pre>
                <p v-else class="hint">點開郵件後才會載入內容；目前沒有可顯示的 body。</p>
                <div class="mail-attachments">
                  <div class="attachment-header">
                    <span>附件</span>
                    <el-tag effect="plain">{{ selectedMailAttachments.length }}</el-tag>
                  </div>
                  <div v-if="isAttachmentListLoading(mail)" class="pane-loading">
                    <span>附件清單載入中...</span>
                  </div>
                  <p v-else-if="selectedMailAttachments.length === 0" class="hint">這封郵件沒有附件。</p>
                  <div v-else class="attachment-list">
                    <div v-for="attachment in selectedMailAttachments" :key="attachment.attachmentId" class="attachment-row">
                      <span class="attachment-main">
                        <strong>{{ attachment.name }}</strong>
                        <span>{{ attachment.contentType || 'unknown' }} · {{ formatAttachmentSize(attachment.size) }}</span>
                      </span>
                      <span class="attachment-actions">
                        <el-button
                          size="small"
                          :loading="isAttachmentExporting(mail, attachment)"
                          :disabled="isAttachmentExporting(mail, attachment)"
                          @click="exportMailAttachment(mail, attachment)"
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
            </article>
            <div v-if="mails.length > 0 && !selectedMail" class="hint">
              選取一封郵件後，內容會展開在 subject 下方。
            </div>
            <div v-if="loadingMails" class="pane-loading">
              <span>Outlook 郵件抓取中...</span>
            </div>
          </div>
        </section>

        <div class="outlook-inspector-column">
          <section class="panel outlook-property-pane">
            <div class="panel-header">
              <div class="panel-title">
                <el-icon><PriceTag /></el-icon>
                <span>新增郵件屬性</span>
              </div>
            </div>

            <div class="inspector-panel-body">
              <div class="library-group">
                <div class="category-heading-row">
                  <div class="library-heading">Master Category List</div>
                  <el-button
                    :icon="Refresh"
                    circle
                    size="small"
                    :loading="loadingCategories"
                    :disabled="outlookBusy && !loadingCategories"
                    @click="requestCategories"
                  />
                </div>
                <div class="category-add-row">
                  <el-input
                    v-model="categoryCreateDraft"
                    :disabled="outlookBusy"
                    placeholder="新增或更新分類名稱"
                    @keydown.enter.prevent="addCategoryToMasterList"
                  />
                  <el-select v-model="categoryCreateColor" class="category-color-select" :disabled="outlookBusy">
                    <el-option
                      v-for="option in categoryColorOptions"
                      :key="option.value"
                      :label="option.label"
                      :value="option.value"
                    >
                      <span class="category-option">
                        <span class="category-swatch" :style="categoryColorStyle(option.value)" />
                        <span>{{ option.label }}</span>
                      </span>
                    </el-option>
                  </el-select>
                  <el-button :loading="operationLoading" :disabled="outlookBusy || !categoryCreateDraft.trim()" @click="addCategoryToMasterList">
                    儲存
                  </el-button>
                </div>
              </div>

              <div class="category-list">
                <div v-if="categories.length === 0 && !loadingCategories" class="hint">尚未取得 Outlook master category list。</div>
                <div v-for="category in categories" :key="category.name" class="category-row">
                  <span class="category-name">
                    <span class="category-swatch" :style="categoryColorStyle(category.color)" />
                    <span>{{ category.name }}</span>
                  </span>
                  <el-select
                    :model-value="category.color || 'olCategoryColorNone'"
                    class="category-row-select"
                    :disabled="outlookBusy"
                    @change="(value: string | number | boolean) => updateCategoryColor(category, String(value))"
                  >
                    <el-option
                      v-for="option in categoryColorOptions"
                      :key="option.value"
                      :label="option.label"
                      :value="option.value"
                    >
                      <span class="category-option">
                        <span class="category-swatch" :style="categoryColorStyle(option.value)" />
                        <span>{{ option.label }}</span>
                      </span>
                    </el-option>
                  </el-select>
                </div>
                <div v-if="loadingCategories" class="pane-loading">
                  <span>Outlook category 同步中...</span>
                </div>
              </div>

              <div class="inspector-note">Flag、日期、已讀狀態與套用哪些 category 都是單封郵件屬性；這裡只管理 Outlook 全域分類名稱與顏色。</div>
            </div>
          </section>

          <section class="panel outlook-property-pane">
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
                  <span>{{ selectedMail.senderName }} &lt;{{ selectedMail.senderEmail }}&gt;</span>
                  <span>{{ formatDateTime(selectedMail.receivedTime) }}</span>
                  <span>來源：{{ selectedMailFolderName }}</span>
                </div>
                <div v-if="!selectedMailHasIdentity" class="identity-warning">
                  這封郵件缺少 id，Add-in 需在 PushMails 回傳 Outlook EntryID 或穩定識別後才能修改或移動。
                </div>

                <div class="marker-tags">
                  <el-tag
                    class="clickable-marker-tag"
                    :class="{ disabled: outlookBusy }"
                    :type="mailPropertiesDraft.isRead ? 'info' : 'warning'"
                    effect="plain"
                    role="button"
                    tabindex="0"
                    :aria-disabled="outlookBusy"
                    @click="!outlookBusy && (mailPropertiesDraft.isRead = !mailPropertiesDraft.isRead)"
                    @keydown.enter.prevent="!outlookBusy && (mailPropertiesDraft.isRead = !mailPropertiesDraft.isRead)"
                    @keydown.space.prevent="!outlookBusy && (mailPropertiesDraft.isRead = !mailPropertiesDraft.isRead)"
                  >
                    {{ mailPropertiesDraft.isRead ? '已讀' : '未讀' }}
                  </el-tag>
                  <el-tag v-if="selectedMail.isMarkedAsTask" type="danger" effect="plain">
                    {{ flagIntervalLabel(selectedMail.flagInterval) }}
                  </el-tag>
                  <el-tag v-if="selectedMail.taskDueDate" type="info" effect="plain">
                    到期 {{ formatDateTime(selectedMail.taskDueDate) }}
                  </el-tag>
                  <el-tag v-if="selectedMail.importance === 'high'" type="danger" effect="plain">高重要性</el-tag>
                  <el-tag
                    v-for="category in selectedMailCategories"
                    :key="category"
                    effect="dark"
                    :style="categoryTagStyle(category)"
                  >
                    {{ category }}
                  </el-tag>
                </div>

                <div class="mail-property-form">
                  <div class="inspector-field">
                    <span>旗標種類</span>
                    <el-select v-model="mailPropertiesDraft.flagInterval" :disabled="outlookBusy">
                      <el-option
                        v-for="option in flagIntervalOptions"
                        :key="option.value"
                        :label="option.label"
                        :value="option.value"
                      />
                    </el-select>
                  </div>

                  <div class="inspector-field">
                    <span>旗標文字</span>
                    <el-input v-model="mailPropertiesDraft.flagRequest" :disabled="outlookBusy || mailPropertiesDraft.flagInterval === 'none'" placeholder="例如：今天" />
                  </div>

                  <div class="date-grid">
                    <div class="inspector-field">
                      <span>自訂開始日</span>
                      <el-date-picker
                        v-model="mailPropertiesDraft.taskStartDate"
                        type="date"
                        value-format="YYYY-MM-DD"
                        :disabled="outlookBusy || mailPropertiesDraft.flagInterval !== 'custom'"
                        placeholder="選擇日期"
                      />
                    </div>
                    <div class="inspector-field">
                      <span>自訂到期日</span>
                      <el-date-picker
                        v-model="mailPropertiesDraft.taskDueDate"
                        type="date"
                        value-format="YYYY-MM-DD"
                        :disabled="outlookBusy || mailPropertiesDraft.flagInterval !== 'custom'"
                        placeholder="選擇日期"
                      />
                    </div>
                  </div>
                  <div v-if="mailPropertiesDraft.flagInterval !== 'custom'" class="field-hint">
                    選擇「自訂日期」後即可設定自訂開始日與到期日。
                  </div>

                  <div class="inspector-field">
                    <span>套用分類</span>
                    <el-select
                      v-model="mailPropertiesDraft.categories"
                      multiple
                      filterable
                      collapse-tags
                      :disabled="outlookBusy"
                      placeholder="選擇分類"
                    >
                      <el-option
                        v-for="category in categories"
                        :key="category.name"
                        :label="category.name"
                        :value="category.name"
                      />
                    </el-select>
                  </div>

                  <div class="inspector-actions commit-actions">
                    <el-button @click="resetMailPropertiesDraft(selectedMail)">重設</el-button>
                    <el-button
                      type="primary"
                      size="large"
                      class="commit-button"
                      :loading="operationLoading"
                      :disabled="!selectedMailHasIdentity || (outlookBusy && !operationLoading)"
                      @click="applyMailProperties(selectedMail)"
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
        </div>

      </main>

      <main v-else-if="activeView === 'chat'" class="chat-layout">
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

      <main v-else-if="activeView === 'calendar'" class="calendar-layout">
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
              <template v-for="week in calendarWeeks" :key="week.map((day) => day.key).join('-')">
                <div
                  v-for="day in week"
                  :key="day.key"
                  class="calendar-day"
                  :class="{ muted: !day.inMonth, today: day.isToday }"
                >
                  <div class="calendar-day-number">{{ day.dayNumber }}</div>
                  <button
                    v-for="event in day.events"
                    :key="event.id || `${event.start}-${event.subject}`"
                    class="calendar-event"
                    type="button"
                    @click="selectCalendarEvent(event)"
                  >
                    <span>{{ formatTime(event.start) }}</span>
                    <strong>{{ event.subject }}</strong>
                  </button>
                </div>
              </template>
            </div>

            <aside class="calendar-detail">
              <template v-if="selectedCalendarEvent">
                <div class="calendar-detail-title">{{ selectedCalendarEvent.subject }}</div>
                <div class="rule-detail">
                  <span>{{ formatDateTime(selectedCalendarEvent.start) }} - {{ formatTime(selectedCalendarEvent.end) }}</span>
                  <span>地點：{{ selectedCalendarEvent.location || '-' }}</span>
                  <span>召集人：{{ selectedCalendarEvent.organizer || '-' }}</span>
                  <span>出席者：{{ selectedCalendarEvent.requiredAttendees || '-' }}</span>
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

      <main v-else-if="activeView === 'admin'" class="admin-layout">
        <section class="panel">
          <div class="panel-header">
            <div class="panel-title">
              <el-icon><Connection /></el-icon>
              <span>Outlook Add-in Status</span>
            </div>
            <div class="admin-actions">
              <el-button :loading="loadingSignalRPing" :disabled="!addinStatus.connected" @click="requestSignalRPing">
                SignalR Ping
              </el-button>
              <el-button :icon="Refresh" @click="refreshAdminData">Refresh</el-button>
            </div>
          </div>

          <div class="status-grid">
            <div class="status-item">
              <span class="status-label">Connection</span>
              <strong :class="addinStatus.connected ? 'online' : 'offline'">
                {{ addinStatus.connected ? 'Online' : 'Offline' }}
              </strong>
            </div>
            <div class="status-item">
              <span class="status-label">Last Connect</span>
              <strong>{{ formatTime(addinStatus.lastPollTime) }}</strong>
            </div>
            <div class="status-item">
              <span class="status-label">Last Push</span>
              <strong>{{ formatTime(addinStatus.lastPushTime) }}</strong>
            </div>
            <div class="status-item">
              <span class="status-label">Last Command</span>
              <strong>{{ addinStatus.lastCommand || '-' }}</strong>
            </div>
          </div>
        </section>

        <section class="panel">
          <div class="panel-header">
            <div class="panel-title">Attachment Export</div>
          </div>

          <div class="admin-settings">
            <div class="status-grid">
              <div class="status-item">
                <span class="status-label">Current Root</span>
                <strong>{{ attachmentExportSettings.rootPath || '載入中...' }}</strong>
              </div>
              <div class="status-item">
                <span class="status-label">Default Root</span>
                <strong>{{ attachmentExportSettings.defaultRootPath || '載入中...' }}</strong>
              </div>
            </div>
            <div class="inspector-field">
              <span>Export root</span>
              <el-input v-model="attachmentExportRootDraft" :placeholder="attachmentExportSettings.defaultRootPath || '$HOME/SmartOffice/Attachments'" />
            </div>
            <div class="field-hint">
              macOS / Linux 預設會放在使用者 home 底下的 SmartOffice/Attachments；Windows 會優先使用 D:\SmartOffice\Attachments，沒有 D 槽時使用 C:\SmartOffice\Attachments。
            </div>
            <div class="admin-actions">
              <el-button :loading="savingAttachmentExportSettings" @click="saveAttachmentExportSettings">
                儲存
              </el-button>
              <el-button :disabled="savingAttachmentExportSettings" @click="resetAttachmentExportRoot">
                使用預設
              </el-button>
            </div>
          </div>
        </section>

        <section class="panel">
          <div class="panel-header">
            <div class="panel-title">Add-in Logs</div>
          </div>

          <div class="logs">
            <p v-if="addinLogs.length === 0">No logs yet.</p>
            <div v-for="(log, index) in addinLogs" :key="`${log.timestamp}-${index}`" class="log-entry" :class="log.level">
              <span>[{{ formatTime(log.timestamp) }}]</span>
              <span>[{{ log.level.toUpperCase() }}]</span>
              <span>{{ log.message }}</span>
            </div>
          </div>
        </section>
      </main>

      <main v-else class="swagger-layout">
        <iframe class="swagger-frame" src="/swagger/index.html" title="Swagger" />
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
    </div>
  </el-config-provider>
</template>
