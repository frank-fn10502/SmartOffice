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
} from '@element-plus/icons-vue'
import FolderNode from './components/FolderNode.vue'
import { useOutlookDashboard } from './composables/useOutlookDashboard'
import type { AppView } from './models/outlook'
import { formatDateTime, formatTime } from './utils/formatters'

const {
  activeView,
  addCategoryToMasterList,
  addinLogs,
  addinStatus,
  applyMailProperties,
  calendarEvents,
  calendarMonthLabel,
  calendarWeekdays,
  calendarWeeks,
  cancelCreateFolder,
  categories,
  categoryColorOptions,
  categoryColorStyle,
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
  fetchMailsFromContext,
  flagIntervalLabel,
  flagIntervalOptions,
  folderContextMenu,
  loadingCalendar,
  loadingFolders,
  loadingMails,
  loadingSignalRPing,
  mailCount,
  mailHtmlSandbox,
  mailPropertiesDraft,
  mailRange,
  mailStats,
  mails,
  moveDraggedMail,
  openFolderContextMenu,
  operationLoading,
  outlookBusy,
  outlookBusyText,
  refreshAdminData,
  requestCalendar,
  requestFolders,
  requestSignalRPing,
  requestMails,
  resetMailPropertiesDraft,
  selectedFolderName,
  selectedFolderPath,
  selectedCalendarEvent,
  selectedMail,
  selectedMailCategories,
  selectedMailHasIdentity,
  selectedMailHtml,
  selectedMailIndex,
  selectedMailIsOpen,
  selectFolder,
  selectCalendarEvent,
  selectMail,
  sendChat,
  goToCurrentCalendarMonth,
  setDragOverFolder,
  signalRState,
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
            <p v-if="visibleFolders.length === 0" class="hint">Waiting for folders...</p>
            <FolderNode
              v-for="folder in visibleFolders"
              :key="folder.folderPath"
              :folder="folder"
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
            <div v-if="outlookBusy" class="pane-loading">
              <span>{{ outlookBusyText }}</span>
            </div>
          </div>
        </section>

        <section class="panel outlook-mail-pane">
          <div class="panel-header">
            <div class="panel-title">
              <el-icon><Document /></el-icon>
              <span>{{ selectedFolderName }}</span>
              <el-tag effect="plain">{{ mails.length }}</el-tag>
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
          </div>

          <div class="mail-table">
            <p v-if="mails.length === 0" class="hint">選取左邊 folder 後抓取郵件。</p>
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
                  <el-tag v-if="mail.categories" type="success" effect="plain">{{ mail.categories }}</el-tag>
                </span>
              </button>

              <div v-if="selectedMailIndex === index && selectedMailIsOpen" class="mail-inline-detail">
                <el-button size="small" @click="selectedMailHtml = !selectedMailHtml">
                  {{ selectedMailHtml ? '切到文字' : '切到 HTML' }}
                </el-button>
                <iframe
                  v-if="selectedMailHtml"
                  class="mail-html"
                  :sandbox="mailHtmlSandbox"
                  referrerpolicy="no-referrer"
                  :srcdoc="mail.bodyHtml || mail.body"
                />
                <pre v-else class="mail-text">{{ mail.body }}</pre>
              </div>
            </article>
            <div v-if="mails.length > 0 && !selectedMail" class="hint">
              選取一封郵件後，內容會展開在 subject 下方。
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
                <div class="library-heading">Master Category List</div>
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
                <div v-if="categories.length === 0" class="hint">尚未取得 Outlook master category list。</div>
                <div v-for="category in categories" :key="category.name" class="category-row">
                  <span class="category-name">
                    <span class="category-swatch" :style="categoryColorStyle(category.color)" />
                    <span>{{ category.name }}</span>
                  </span>
                  <el-select
                    :model-value="category.color || 'preset0'"
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
                  <span>來源：{{ selectedFolderName }}</span>
                </div>
                <div v-if="!selectedMailHasIdentity" class="identity-warning">
                  這封郵件缺少 id，Add-in 需在 PushMails 回傳 Outlook EntryID 或穩定識別後才能修改或移動。
                </div>

                <div class="marker-tags">
                  <el-tag :type="selectedMail.isRead ? 'info' : 'warning'" effect="plain">
                    {{ selectedMail.isRead ? '已讀' : '未讀' }}
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
                    type="success"
                    effect="plain"
                  >
                    {{ category }}
                  </el-tag>
                </div>

                <div class="mail-property-form">
                  <label class="inspector-check">
                    <el-switch v-model="mailPropertiesDraft.isRead" :disabled="outlookBusy" />
                    <span>{{ mailPropertiesDraft.isRead ? '已讀' : '未讀' }}</span>
                  </label>

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
