<script setup lang="ts">
import { computed, nextTick, onMounted, ref } from 'vue'
import * as signalR from '@microsoft/signalr'
import {
  Calendar,
  ChatDotRound,
  Connection,
  Document,
  Flag,
  Folder,
  Monitor,
  PriceTag,
  Refresh,
  Tickets,
} from '@element-plus/icons-vue'
import FolderNode from './components/FolderNode.vue'

interface FolderDto {
  name: string
  folderPath: string
  itemCount: number
  subFolders: FolderDto[]
}

interface MailItemDto {
  id: string
  subject: string
  senderName: string
  senderEmail: string
  receivedTime: string
  body: string
  bodyHtml: string
  folderPath: string
  categories: string
  isRead: boolean
  isMarkedAsTask: boolean
  importance: string
  sensitivity: string
}

interface OutlookRuleDto {
  name: string
  enabled: boolean
  executionOrder: number
  ruleType: string
  conditions: string[]
  actions: string[]
  exceptions: string[]
}

interface CalendarEventDto {
  id: string
  subject: string
  start: string
  end: string
  location: string
  organizer: string
  requiredAttendees: string
  isRecurring: boolean
  busyStatus: string
}

interface ChatMessageDto {
  id?: string
  source: string
  text: string
  timestamp: string
}

interface AddinStatusDto {
  connected: boolean
  lastPollTime?: string
  lastPushTime?: string
  lastCommand: string
}

interface AddinLogEntry {
  level: 'info' | 'warn' | 'error' | string
  message: string
  timestamp: string
}

type AppView = 'normal' | 'outlook' | 'admin' | 'swagger'

const activeView = ref<AppView>('normal')
const signalRState = ref<'connected' | 'reconnecting' | 'disconnected'>('disconnected')
const folders = ref<FolderDto[]>([])
const mails = ref<MailItemDto[]>([])
const rules = ref<OutlookRuleDto[]>([])
const calendarEvents = ref<CalendarEventDto[]>([])
const chatMessages = ref<ChatMessageDto[]>([])
const addinStatus = ref<AddinStatusDto>({
  connected: false,
  lastCommand: '',
})
const addinLogs = ref<AddinLogEntry[]>([])
const selectedFolderPath = ref('')
const expandedFolders = ref<Set<string>>(new Set())
const openMailIndexes = ref<Set<number>>(new Set())
const htmlMailIndexes = ref<Set<number>>(new Set())
const mailRange = ref('1d')
const mailCount = ref(10)
const chatText = ref('')
const loadingFolders = ref(false)
const loadingMails = ref(false)
const loadingRules = ref(false)
const loadingCalendar = ref(false)
const chatPanelRef = ref<HTMLElement | null>(null)
const mailHtmlSandbox = 'allow-same-origin allow-popups allow-popups-to-escape-sandbox'

const mailStats = computed(() => ({
  unread: mails.value.filter((mail) => !mail.isRead).length,
  flagged: mails.value.filter((mail) => mail.isMarkedAsTask).length,
  highImportance: mails.value.filter((mail) => mail.importance === 'high').length,
  categorized: mails.value.filter((mail) => Boolean(mail.categories)).length,
}))

const visibleFolders = computed(() => {
  return folders.value.flatMap((root) => {
    if (root.subFolders?.length) return root.subFolders.filter((folder) => !isHiddenFolder(folder.name))
    return isHiddenFolder(root.name) ? [] : [root]
  })
})

function isHiddenFolder(name: string) {
  const hiddenNames = [
    'common views',
    'finder',
    'reminders',
    'quick step',
    'conversation history',
    'conversation action',
    'server failures',
    'local failures',
    'conflicts',
    'sync issues',
    'rss',
    'social network',
    'people',
    'externalcontacts',
    'yammer',
  ]
  const lowerName = name.toLowerCase()
  return hiddenNames.some((hidden) => lowerName.includes(hidden))
}

function visibleChildren(folder: FolderDto) {
  return (folder.subFolders ?? []).filter((child) => !isHiddenFolder(child.name))
}

function folderType(name: string) {
  const lowerName = name.toLowerCase()
  if (lowerName === 'inbox') return 'inbox'
  if (lowerName === 'sent items' || lowerName.includes('sent')) return 'sent'
  if (lowerName === 'drafts') return 'drafts'
  if (lowerName === 'deleted items' || lowerName.includes('deleted')) return 'deleted'
  if (lowerName === 'junk email' || lowerName === 'junk e-mail') return 'junk'
  if (lowerName === 'archive') return 'archive'
  if (lowerName === 'outbox') return 'outbox'
  return 'normal'
}

function folderIcon(name: string) {
  const icons: Record<string, string> = {
    inbox: 'Inbox',
    sent: 'Sent',
    drafts: 'Draft',
    deleted: 'Trash',
    junk: 'Junk',
    archive: 'Archive',
    outbox: 'Out',
    normal: 'Folder',
  }
  return icons[folderType(name)]
}

function formatTime(value?: string) {
  if (!value) return '-'
  return new Date(value).toLocaleTimeString()
}

function formatDateTime(value: string) {
  return new Date(value).toLocaleString()
}

function pollUntil(check: () => Promise<boolean>, timeoutMs: number) {
  return new Promise<boolean>((resolve) => {
    const start = Date.now()
    const timer = window.setInterval(async () => {
      try {
        const done = await check()
        if (done || Date.now() - start >= timeoutMs) {
          window.clearInterval(timer)
          resolve(done)
        }
      } catch {
        if (Date.now() - start >= timeoutMs) {
          window.clearInterval(timer)
          resolve(false)
        }
      }
    }, 1200)
  })
}

async function getJson<T>(url: string): Promise<T> {
  const response = await fetch(url)
  if (!response.ok) throw new Error(`Request failed: ${response.status}`)
  return response.json() as Promise<T>
}

async function postJson<T>(url: string, body?: unknown): Promise<T> {
  const response = await fetch(url, {
    method: 'POST',
    headers: body ? { 'Content-Type': 'application/json' } : undefined,
    body: body ? JSON.stringify(body) : undefined,
  })
  if (!response.ok) throw new Error(`Request failed: ${response.status}`)
  return response.json() as Promise<T>
}

async function loadCachedFolders() {
  folders.value = await getJson<FolderDto[]>('/api/outlook/folders')
  selectDefaultFolder()
}

async function requestFolders() {
  loadingFolders.value = true
  try {
    await postJson('/api/outlook/request-folders')
    await pollUntil(async () => {
      await loadCachedFolders()
      return folders.value.length > 0
    }, 30000)
  } finally {
    loadingFolders.value = false
  }
}

function selectDefaultFolder() {
  if (selectedFolderPath.value || visibleFolders.value.length === 0) return
  const inbox = visibleFolders.value.find((folder) => folderType(folder.name) === 'inbox')
  selectedFolderPath.value = inbox?.folderPath ?? visibleFolders.value[0]?.folderPath ?? ''
}

function toggleFolder(path: string) {
  const next = new Set(expandedFolders.value)
  if (next.has(path)) next.delete(path)
  else next.add(path)
  expandedFolders.value = next
}

function selectFolder(path: string) {
  selectedFolderPath.value = path
}

async function loadCachedMails() {
  mails.value = await getJson<MailItemDto[]>('/api/outlook/mails')
}

async function loadCachedRules() {
  rules.value = await getJson<OutlookRuleDto[]>('/api/outlook/rules')
}

async function loadCachedCalendar() {
  calendarEvents.value = await getJson<CalendarEventDto[]>('/api/outlook/calendar')
}

async function requestMails() {
  loadingMails.value = true
  openMailIndexes.value = new Set()
  htmlMailIndexes.value = new Set()
  try {
    await postJson('/api/outlook/request-mails', {
      folderPath: selectedFolderPath.value,
      range: mailRange.value,
      maxCount: mailCount.value,
    })
    await pollUntil(async () => {
      await loadCachedMails()
      return mails.value.length > 0
    }, 30000)
  } finally {
    loadingMails.value = false
  }
}

async function requestRules() {
  loadingRules.value = true
  try {
    await postJson('/api/outlook/request-rules')
    await pollUntil(async () => {
      await loadCachedRules()
      return rules.value.length > 0
    }, 30000)
  } finally {
    loadingRules.value = false
  }
}

async function requestCalendar() {
  loadingCalendar.value = true
  try {
    await postJson('/api/outlook/request-calendar', { daysForward: 14 })
    await pollUntil(async () => {
      await loadCachedCalendar()
      return calendarEvents.value.length > 0
    }, 30000)
  } finally {
    loadingCalendar.value = false
  }
}

function toggleMail(index: number) {
  const next = new Set(openMailIndexes.value)
  if (next.has(index)) next.delete(index)
  else next.add(index)
  openMailIndexes.value = next
}

function toggleMailFormat(index: number) {
  const next = new Set(htmlMailIndexes.value)
  if (next.has(index)) next.delete(index)
  else next.add(index)
  htmlMailIndexes.value = next
}

async function loadChat() {
  chatMessages.value = await getJson<ChatMessageDto[]>('/api/outlook/chat')
  await scrollChatToBottom()
}

async function sendChat() {
  const text = chatText.value.trim()
  if (!text) return
  chatText.value = ''
  await postJson('/api/outlook/chat', { source: 'web', text })
}

async function refreshAdminData() {
  const [status, logs] = await Promise.all([
    getJson<AddinStatusDto>('/api/outlook/admin/status'),
    getJson<AddinLogEntry[]>('/api/outlook/admin/logs'),
  ])
  addinStatus.value = status
  addinLogs.value = logs
}

async function switchView(view: AppView) {
  activeView.value = view
  if (view === 'admin') await refreshAdminData()
}

async function scrollChatToBottom() {
  await nextTick()
  if (chatPanelRef.value) chatPanelRef.value.scrollTop = chatPanelRef.value.scrollHeight
}

async function connectSignalR() {
  const connection = new signalR.HubConnectionBuilder()
    .withUrl('/hub/notifications')
    .withAutomaticReconnect()
    .build()

  connection.onreconnecting(() => {
    signalRState.value = 'reconnecting'
  })
  connection.onreconnected(() => {
    signalRState.value = 'connected'
  })
  connection.onclose(() => {
    signalRState.value = 'disconnected'
  })
  connection.on('FoldersUpdated', (items: FolderDto[]) => {
    folders.value = items
    selectDefaultFolder()
    loadingFolders.value = false
  })
  connection.on('MailsUpdated', (items: MailItemDto[]) => {
    mails.value = items
    loadingMails.value = false
  })
  connection.on('RulesUpdated', (items: OutlookRuleDto[]) => {
    rules.value = items
    loadingRules.value = false
  })
  connection.on('CalendarUpdated', (items: CalendarEventDto[]) => {
    calendarEvents.value = items
    loadingCalendar.value = false
  })
  connection.on('NewChatMessage', async (message: ChatMessageDto) => {
    chatMessages.value = [...chatMessages.value, message]
    await scrollChatToBottom()
  })
  connection.on('AddinStatus', (status: AddinStatusDto) => {
    addinStatus.value = status
  })
  connection.on('AddinLog', (logs: AddinLogEntry[]) => {
    addinLogs.value = logs
  })

  try {
    await connection.start()
    signalRState.value = 'connected'
  } catch {
    signalRState.value = 'disconnected'
  }
}

onMounted(async () => {
  await connectSignalR()
  await Promise.allSettled([
    loadCachedFolders(),
    loadCachedMails(),
    loadCachedRules(),
    loadCachedCalendar(),
    loadChat(),
    refreshAdminData(),
  ])
  if (folders.value.length === 0) await requestFolders()
})
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
              { label: 'Normal', value: 'normal' },
              { label: 'Outlook', value: 'outlook' },
              { label: 'Admin', value: 'admin' },
              { label: 'Swagger', value: 'swagger' },
            ]"
            @update:model-value="(value: string | number | boolean) => switchView(value as AppView)"
          />
        </nav>
      </header>

      <main v-if="activeView === 'normal'" class="normal-layout">
        <section class="panel folder-panel">
          <div class="panel-header">
            <div class="panel-title">
              <el-icon><Folder /></el-icon>
              <span>Folders</span>
            </div>
            <el-button :icon="Refresh" circle :loading="loadingFolders" @click="requestFolders" />
          </div>

          <div class="folder-list">
            <p v-if="visibleFolders.length === 0" class="hint">Waiting for folders...</p>
            <FolderNode
              v-for="folder in visibleFolders"
              :key="folder.folderPath"
              :folder="folder"
              :level="0"
              :expanded-folders="expandedFolders"
              :selected-folder-path="selectedFolderPath"
              @toggle="toggleFolder"
              @select="selectFolder"
            />
          </div>
        </section>

        <section class="panel mail-panel">
          <div class="panel-header">
            <div class="panel-title">
              <el-icon><Document /></el-icon>
              <span>Mails</span>
              <el-tag effect="plain">{{ mails.length }}</el-tag>
            </div>

            <div class="mail-controls">
              <el-select v-model="mailRange" class="range-select">
                <el-option label="Today" value="1d" />
                <el-option label="Last 7 days" value="1w" />
                <el-option label="Last 30 days" value="1m" />
              </el-select>
              <el-select v-model="mailCount" class="count-select">
                <el-option :value="10" label="10" />
                <el-option :value="20" label="20" />
                <el-option :value="30" label="30" />
                <el-option :value="100" label="100" />
              </el-select>
              <el-button type="primary" :loading="loadingMails" @click="requestMails">Fetch Mails</el-button>
            </div>
          </div>

          <div class="mail-list">
            <p v-if="mails.length === 0" class="hint">Click Fetch Mails to load emails from the selected folder.</p>
            <article v-for="(mail, index) in mails" :key="`${mail.receivedTime}-${index}`" class="mail-item">
              <button class="mail-summary" type="button" @click="toggleMail(index)">
                <span class="mail-main">
                  <span class="mail-subject">{{ mail.subject }}</span>
                  <span class="mail-sender">{{ mail.senderName }} &lt;{{ mail.senderEmail }}&gt;</span>
                </span>
                <span class="mail-time">{{ formatDateTime(mail.receivedTime) }}</span>
              </button>

              <div v-if="openMailIndexes.has(index)" class="mail-detail">
                <el-button size="small" @click="toggleMailFormat(index)">
                  {{ htmlMailIndexes.has(index) ? 'Switch to Text' : 'Switch to HTML' }}
                </el-button>
                <iframe
                  v-if="htmlMailIndexes.has(index)"
                  class="mail-html"
                  :sandbox="mailHtmlSandbox"
                  referrerpolicy="no-referrer"
                  :srcdoc="mail.bodyHtml || mail.body"
                />
                <pre v-else class="mail-text">{{ mail.body }}</pre>
              </div>
            </article>
          </div>
        </section>

        <section class="panel chat-panel">
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

      <main v-else-if="activeView === 'outlook'" class="outlook-layout">
        <section class="panel">
          <div class="panel-header">
            <div class="panel-title">
              <el-icon><PriceTag /></el-icon>
              <span>Mail Markers</span>
            </div>
            <el-button :icon="Refresh" :loading="loadingMails" @click="requestMails">Refresh Mails</el-button>
          </div>

          <div class="status-grid">
            <div class="status-item">
              <span class="status-label">Unread</span>
              <strong>{{ mailStats.unread }}</strong>
            </div>
            <div class="status-item">
              <span class="status-label">Flagged</span>
              <strong>{{ mailStats.flagged }}</strong>
            </div>
            <div class="status-item">
              <span class="status-label">High Importance</span>
              <strong>{{ mailStats.highImportance }}</strong>
            </div>
            <div class="status-item">
              <span class="status-label">Categorized</span>
              <strong>{{ mailStats.categorized }}</strong>
            </div>
          </div>

          <div class="workspace-list">
            <p v-if="mails.length === 0" class="hint">No cached mails yet.</p>
            <article v-for="(mail, index) in mails" :key="mail.id || `${mail.receivedTime}-${index}`" class="workspace-item">
              <div class="workspace-item-main">
                <strong>{{ mail.subject }}</strong>
                <span>{{ mail.senderName }} · {{ formatDateTime(mail.receivedTime) }}</span>
              </div>
              <div class="marker-tags">
                <el-tag :type="mail.isRead ? 'info' : 'warning'" effect="plain">
                  {{ mail.isRead ? 'Read' : 'Unread' }}
                </el-tag>
                <el-tag v-if="mail.isMarkedAsTask" type="danger" effect="plain">
                  <el-icon><Flag /></el-icon>
                  Flag
                </el-tag>
                <el-tag v-if="mail.importance === 'high'" type="danger" effect="plain">High</el-tag>
                <el-tag v-if="mail.categories" type="success" effect="plain">{{ mail.categories }}</el-tag>
              </div>
            </article>
          </div>
        </section>

        <section class="panel">
          <div class="panel-header">
            <div class="panel-title">
              <el-icon><Tickets /></el-icon>
              <span>Rules</span>
              <el-tag effect="plain">{{ rules.length }}</el-tag>
            </div>
            <el-button :icon="Refresh" :loading="loadingRules" @click="requestRules">Fetch Rules</el-button>
          </div>

          <div class="workspace-list">
            <p v-if="rules.length === 0" class="hint">No cached rules yet.</p>
            <article v-for="rule in rules" :key="`${rule.executionOrder}-${rule.name}`" class="workspace-item">
              <div class="workspace-item-main">
                <strong>{{ rule.executionOrder }}. {{ rule.name }}</strong>
                <span>{{ rule.ruleType }} · {{ rule.enabled ? 'Enabled' : 'Disabled' }}</span>
              </div>
              <div class="rule-detail">
                <span>When: {{ rule.conditions.join('; ') || '-' }}</span>
                <span>Do: {{ rule.actions.join('; ') || '-' }}</span>
                <span v-if="rule.exceptions.length">Except: {{ rule.exceptions.join('; ') }}</span>
              </div>
            </article>
          </div>
        </section>

        <section class="panel">
          <div class="panel-header">
            <div class="panel-title">
              <el-icon><Calendar /></el-icon>
              <span>Calendar</span>
              <el-tag effect="plain">{{ calendarEvents.length }}</el-tag>
            </div>
            <el-button :icon="Refresh" :loading="loadingCalendar" @click="requestCalendar">Fetch Calendar</el-button>
          </div>

          <div class="workspace-list">
            <p v-if="calendarEvents.length === 0" class="hint">No cached calendar events yet.</p>
            <article v-for="event in calendarEvents" :key="event.id || `${event.start}-${event.subject}`" class="workspace-item">
              <div class="workspace-item-main">
                <strong>{{ event.subject }}</strong>
                <span>{{ formatDateTime(event.start) }} - {{ formatTime(event.end) }}</span>
              </div>
              <div class="rule-detail">
                <span>Location: {{ event.location || '-' }}</span>
                <span>Organizer: {{ event.organizer || '-' }}</span>
                <span>Attendees: {{ event.requiredAttendees || '-' }}</span>
              </div>
              <div class="marker-tags">
                <el-tag effect="plain">{{ event.busyStatus || 'unknown' }}</el-tag>
                <el-tag v-if="event.isRecurring" type="warning" effect="plain">Recurring</el-tag>
              </div>
            </article>
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
            <el-button :icon="Refresh" @click="refreshAdminData">Refresh</el-button>
          </div>

          <div class="status-grid">
            <div class="status-item">
              <span class="status-label">Connection</span>
              <strong :class="addinStatus.connected ? 'online' : 'offline'">
                {{ addinStatus.connected ? 'Online' : 'Offline' }}
              </strong>
            </div>
            <div class="status-item">
              <span class="status-label">Last Poll</span>
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
    </div>
  </el-config-provider>
</template>
