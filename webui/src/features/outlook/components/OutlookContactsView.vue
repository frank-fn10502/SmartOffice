<script setup lang="ts">
import { computed, onMounted, ref, watch } from 'vue'
import { Refresh, Search, UserFilled } from '@element-plus/icons-vue'
import { ElMessage } from 'element-plus'
import { normalizeAddressBookContact, outlookApi } from '../api/outlook'
import type { AddressBookContactDto } from '../models/outlook'
import { fetchResultEndpoint, requestIdFromResponse } from '../composables/outlookRequests'
import { formatDateTime } from '../utils/formatters'
import type { OutlookDashboardState } from '../composables/useOutlookDashboard'

const props = defineProps<{
  dashboard?: OutlookDashboardState
}>()

const contacts = ref<AddressBookContactDto[]>([])
const selectedContact = ref<AddressBookContactDto | null>(null)
const query = ref('')
const loadingContacts = ref(false)
const lookupLoading = ref(false)
const syncing = ref(false)
const lastUpdatedAt = ref<Date | null>(null)
const loadMessage = ref('尚未載入通訊錄。')
const lookupEmail = ref('')
const lookupMessage = ref('')
let initialLoadRequested = false

const knownCount = computed(() => contacts.value.filter((contact) => contact.isKnown).length)
const selfCount = computed(() => contacts.value.filter((contact) => contact.isLikelySelf).length)
const groupCount = computed(() => contacts.value.filter((contact) => contact.isGroup).length)
const personCount = computed(() => contacts.value.filter((contact) => !contact.isGroup).length)
const lastUpdatedText = computed(() => lastUpdatedAt.value
  ? lastUpdatedAt.value.toLocaleString('zh-TW', { hour12: false })
  : '尚未完成同步')

function contactTitle(contact: AddressBookContactDto) {
  return contact.displayName || contact.smtpAddress || '(unknown)'
}

function contactKey(contact: AddressBookContactDto) {
  return (contact.smtpAddress || contact.rawAddress || contact.id || contact.displayName).trim().toLowerCase()
}

function isSelectedContact(contact: AddressBookContactDto) {
  return Boolean(selectedContact.value && contactKey(selectedContact.value) === contactKey(contact))
}

function relationLabel(kind: string) {
  const labels: Record<string, string> = {
    attendee: '出席者',
    bcc: '密件副本',
    cc: '副本',
    group_member: '群組成員',
    organizer: '召集人',
    sender: '寄件者',
    to: '收件者',
  }
  return labels[kind] ?? kind
}

async function loadContacts() {
  if (loadingContacts.value) return
  loadingContacts.value = true
  loadMessage.value = '正在建立 Outlook 通訊錄 request...'
  try {
    const response = await outlookApi.requestAddressBook({
      includeOutlookContacts: true,
      includeAddressLists: true,
      maxContacts: 1000,
      maxAddressEntriesPerList: 500,
      maxGroupMembers: 50,
      maxGroupDepth: 1,
    })
    contacts.value = []
    selectedContact.value = null
    loadMessage.value = 'Outlook 正在分段讀取 Contacts / AddressLists / group metadata...'
    await streamContactsFromRequest(response)
    if (!selectedContact.value || !contacts.value.some((contact) => isSelectedContact(contact))) {
      selectedContact.value = contacts.value[0] ?? null
    }
    lastUpdatedAt.value = new Date()
    loadMessage.value = `通訊錄同步完成，共 ${contacts.value.length} 筆。`
  } catch (error) {
    loadMessage.value = error instanceof Error ? error.message : '通訊錄同步失敗。'
    throw error
  } finally {
    loadingContacts.value = false
  }
}

async function streamContactsFromRequest(response: { requestId?: string; request?: string; data?: unknown }) {
  const requestId = requestIdFromResponse(response)
  if (!requestId) return
  const endpoint = fetchResultEndpoint(response)
  const started = Date.now()
  const timeoutMs = 120000
  const seen = new Map<string, AddressBookContactDto>()
  let cursor = ''
  let fetchedCount = 0

  while (Date.now() - started < timeoutMs) {
    const state = await fetchAddressBookPage(endpoint, requestId, cursor)
    if (!state) {
      await new Promise((resolve) => window.setTimeout(resolve, 500))
      continue
    }
    const rawContacts = Array.isArray(state.data?.contacts) ? state.data.contacts : []
    const pageContacts = rawContacts.map(normalizeAddressBookContact)
    fetchedCount += rawContacts.length
    mergeContacts(seen, pageContacts)

    if (pageContacts.length > 0) {
      contacts.value = filterContacts(Array.from(seen.values()))
      if (!selectedContact.value) selectedContact.value = contacts.value[0] ?? null
    }

    loadMessage.value = state.state === 'completed'
      ? `通訊錄同步完成，共 ${seen.size} 筆。`
      : `通訊錄載入中，已取得 ${seen.size} 筆...`

    if (state.next.hasMore) {
      cursor = state.next.cursor
      continue
    }
    if (state.state === 'completed') return
    if (state.state && !['accepted', 'running'].includes(state.state)) {
      throw new Error(state.message || '通訊錄同步失敗。')
    }

    cursor = fetchedCount > 0 ? String(fetchedCount) : ''
    await new Promise((resolve) => window.setTimeout(resolve, 500))
  }

  throw new Error('通訊錄同步逾時。')
}

async function fetchAddressBookPage(endpoint: string, requestId: string, cursor: string) {
  try {
    return await outlookApi.fetchResult<{ contacts?: unknown[] }>(endpoint, { requestId, cursor, take: 100 })
  } catch (error) {
    if (error instanceof Error && error.message === 'Request failed: 404') return null
    throw error
  }
}

function mergeContacts(target: Map<string, AddressBookContactDto>, nextContacts: AddressBookContactDto[]) {
  for (const contact of nextContacts) {
    const key = contact.smtpAddress || contact.rawAddress || contact.displayName || contact.id
    if (key) target.set(key.toLowerCase(), contact)
  }
}

function filterContacts(source: AddressBookContactDto[]) {
  const text = query.value.trim().toLowerCase()
  return source.filter((contact) => !text
    || contact.displayName.toLowerCase().includes(text)
    || contact.smtpAddress.toLowerCase().includes(text)
    || contact.domain.toLowerCase().includes(text))
}

function requestInitialLoad(reason: 'active-view' | 'background') {
  if (initialLoadRequested || loadingContacts.value || contacts.value.length > 0) return
  if (reason === 'background' && !props.dashboard?.outlookFirstLoadCompleted.value) return
  if (reason === 'active-view' && props.dashboard?.activeView.value !== 'contacts') return
  initialLoadRequested = true
  void loadContacts().catch(() => { initialLoadRequested = false })
}

async function lookupContact() {
  if (!lookupEmail.value.trim()) {
    lookupMessage.value = ''
    return
  }

  lookupLoading.value = true
  try {
    const response = await outlookApi.lookupAddressBookContact(lookupEmail.value)
    lookupMessage.value = response.message
    if (response.contact) {
      selectedContact.value = response.contact
      if (!contacts.value.some((contact) => contact.smtpAddress === response.contact?.smtpAddress)) {
        contacts.value = [response.contact, ...contacts.value]
      }
      ElMessage.success('找到已知關聯')
    } else {
      ElMessage.warning('目前沒有這個 email 的關聯')
    }
  } finally {
    lookupLoading.value = false
  }
}

async function syncAddressBook() {
  if (loadingContacts.value || syncing.value) return
  syncing.value = true
  try {
    await loadContacts()
    ElMessage.success('通訊錄同步完成')
  } finally {
    syncing.value = false
  }
}

onMounted(() => {
  if (!props.dashboard) void loadContacts()
  else requestInitialLoad('active-view')
})

watch(
  () => props.dashboard?.activeView.value,
  () => requestInitialLoad('active-view'),
)

watch(
  () => props.dashboard?.outlookFirstLoadCompleted.value,
  (completed) => {
    if (completed) window.setTimeout(() => requestInitialLoad('background'), 1200)
  },
  { immediate: true },
)
</script>

<template>
  <main class="contacts-layout">
    <section class="panel">
      <div class="panel-header">
        <div class="panel-title">
          <el-icon><UserFilled /></el-icon>
          <span>通訊錄</span>
          <el-tag effect="plain">{{ contacts.length }}</el-tag>
        </div>
        <div class="contacts-actions">
          <el-input
            v-model="query"
            class="contacts-query"
            clearable
            :disabled="loadingContacts"
            :prefix-icon="Search"
            placeholder="姓名、email 或 domain"
            @keyup.enter="loadContacts"
            @clear="loadContacts"
          />
          <el-button :icon="Refresh" :loading="loadingContacts" :disabled="loadingContacts" @click="loadContacts">重新整理</el-button>
          <el-button type="primary" :loading="syncing" :disabled="loadingContacts || syncing" @click="syncAddressBook">同步 Outlook 通訊錄</el-button>
        </div>
      </div>

      <div class="contacts-sync-status" :class="{ loading: loadingContacts }">
        <div>
          <strong>{{ loadingContacts ? '通訊錄載入中...' : '通訊錄狀態' }}</strong>
          <span>{{ loadMessage }}</span>
        </div>
        <div class="contacts-sync-stats">
          <span>{{ personCount }} 個人</span>
          <span>{{ groupCount }} group</span>
          <span>最後更新：{{ lastUpdatedText }}</span>
        </div>
      </div>

      <div class="contacts-page">
        <div class="contacts-list">
          <div class="contacts-stats">
            <span>個人 {{ personCount }}</span>
            <span>Group {{ groupCount }}</span>
            <span>已知關聯 {{ knownCount }}</span>
            <span>可能是自己 {{ selfCount }}</span>
          </div>

          <div v-if="loadingContacts" class="contacts-loading" role="status">
            <span />
            <strong>正在載入通訊錄...</strong>
          </div>

          <button
            v-for="contact in contacts"
            :key="contactKey(contact)"
            class="contact-row"
            :class="{ active: isSelectedContact(contact) }"
            type="button"
            @click="selectedContact = contact"
          >
            <strong>{{ contactTitle(contact) }}</strong>
            <span>{{ contact.smtpAddress || '-' }}</span>
            <small>{{ contact.mailCount }} mails / {{ contact.calendarCount }} calendar</small>
          </button>

          <el-empty v-if="!loadingContacts && contacts.length === 0" description="目前沒有符合條件的聯絡人" />
        </div>

        <aside class="contact-detail">
          <div class="lookup-card">
            <el-input
              v-model="lookupEmail"
              clearable
              placeholder="檢查收件者 email"
              @keyup.enter="lookupContact"
            >
              <template #append>
                <el-button :icon="Search" :loading="lookupLoading" :disabled="lookupLoading" @click="lookupContact" />
              </template>
            </el-input>
            <span v-if="lookupMessage">{{ lookupMessage }}</span>
          </div>

          <template v-if="selectedContact">
            <div class="contact-detail-title">
              <strong>{{ contactTitle(selectedContact) }}</strong>
              <el-tag v-if="selectedContact.isLikelySelf" type="warning" effect="plain">自己</el-tag>
              <el-tag v-else-if="selectedContact.isGroup" type="info" effect="plain">群組</el-tag>
              <el-tag v-else type="success" effect="plain">已知關聯</el-tag>
            </div>

            <div class="rule-detail">
              <span>Email：{{ selectedContact.smtpAddress || '-' }}</span>
              <span>Domain：{{ selectedContact.domain || '-' }}</span>
              <span>來源：{{ selectedContact.sources.join(', ') || selectedContact.source || '-' }}</span>
              <span>公司 / 部門：{{ selectedContact.companyName || '-' }} / {{ selectedContact.department || '-' }}</span>
              <span>職稱：{{ selectedContact.jobTitle || '-' }}</span>
              <span v-if="selectedContact.isGroup">成員數：{{ selectedContact.memberCount }}</span>
              <span v-if="selectedContact.memberOfGroupSmtpAddresses.length > 0">
                隸屬群組：{{ selectedContact.memberOfGroupSmtpAddresses.slice(0, 5).join(', ') }}
              </span>
              <span>最近互動：{{ selectedContact.lastSeen ? formatDateTime(selectedContact.lastSeen) : '-' }}</span>
              <span>最早出現：{{ selectedContact.firstSeen ? formatDateTime(selectedContact.firstSeen) : '-' }}</span>
              <span>關聯分數：{{ selectedContact.relationScore }}</span>
            </div>

            <div class="marker-tags">
              <el-tag v-for="kind in selectedContact.relationKinds" :key="kind" effect="plain">
                {{ relationLabel(kind) }}
              </el-tag>
            </div>

            <div class="contact-evidence">
              <strong>近期依據</strong>
              <span v-if="selectedContact.memberSmtpAddresses.length > 0">
                成員：{{ selectedContact.memberSmtpAddresses.slice(0, 8).join(', ') }}
              </span>
              <span v-if="selectedContact.memberGroupSmtpAddresses.length > 0">
                子群組：{{ selectedContact.memberGroupSmtpAddresses.slice(0, 8).join(', ') }}
              </span>
              <span v-for="subject in selectedContact.sampleSubjects" :key="subject">{{ subject }}</span>
              <span v-if="selectedContact.sampleSubjects.length === 0">目前只有 calendar 或群組 metadata。</span>
            </div>
          </template>

          <div v-else class="empty-inspector">
            選取聯絡人查看關聯來源。
          </div>
        </aside>
      </div>
    </section>
  </main>
</template>
