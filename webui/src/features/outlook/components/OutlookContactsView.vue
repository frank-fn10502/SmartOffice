<script setup lang="ts">
import { computed, onMounted, ref, watch } from 'vue'
import { Refresh, Search, UserFilled } from '@element-plus/icons-vue'
import { ElMessage } from 'element-plus'
import { outlookApi } from '../api/outlook'
import type { AddressBookContactDto } from '../models/outlook'
import { collectOutlookRequestData, waitForOutlookRequest } from '../composables/outlookRequests'
import { formatDateTime } from '../utils/formatters'
import type { OutlookDashboardState } from '../composables/useOutlookDashboard'

const props = defineProps<{
  dashboard?: OutlookDashboardState
}>()

const contacts = ref<AddressBookContactDto[]>([])
const selectedContact = ref<AddressBookContactDto | null>(null)
const query = ref('')
const loading = ref(false)
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
  if (loading.value) return
  loading.value = true
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
    loadMessage.value = 'Outlook 正在讀取 Contacts / AddressLists / group metadata...'
    await waitForOutlookRequest(response, { timeoutMs: 120000 })
    loadMessage.value = '正在分段讀取 Hub fetch-result-address-book...'
    const pages = await collectOutlookRequestData<{ contacts?: AddressBookContactDto[] }>(response)
    const text = query.value.trim().toLowerCase()
    contacts.value = pages
      .flatMap((page) => page.data?.contacts ?? [])
      .filter((contact) => !text
        || contact.displayName.toLowerCase().includes(text)
        || contact.smtpAddress.toLowerCase().includes(text)
        || contact.domain.toLowerCase().includes(text))
    if (!selectedContact.value || !contacts.value.some((contact) => contact.smtpAddress === selectedContact.value?.smtpAddress)) {
      selectedContact.value = contacts.value[0] ?? null
    }
    lastUpdatedAt.value = new Date()
    loadMessage.value = `通訊錄同步完成，共 ${contacts.value.length} 筆。`
  } catch (error) {
    loadMessage.value = error instanceof Error ? error.message : '通訊錄同步失敗。'
    throw error
  } finally {
    loading.value = false
  }
}

function requestInitialLoad(reason: 'active-view' | 'background') {
  if (initialLoadRequested || loading.value || contacts.value.length > 0) return
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

  loading.value = true
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
    loading.value = false
  }
}

async function syncAddressBook() {
  if (loading.value || syncing.value) return
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
            :disabled="loading"
            :prefix-icon="Search"
            placeholder="姓名、email 或 domain"
            @keyup.enter="loadContacts"
            @clear="loadContacts"
          />
          <el-button :icon="Refresh" :loading="loading" :disabled="loading" @click="loadContacts">重新整理</el-button>
          <el-button type="primary" :loading="syncing" :disabled="loading || syncing" @click="syncAddressBook">同步 Outlook 通訊錄</el-button>
        </div>
      </div>

      <div class="contacts-sync-status" :class="{ loading }">
        <div>
          <strong>{{ loading ? '通訊錄載入中...' : '通訊錄狀態' }}</strong>
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

          <div v-if="loading" class="pane-loading" role="status">
            <span />
            <strong>正在載入通訊錄...</strong>
          </div>

          <button
            v-for="contact in contacts"
            :key="contact.smtpAddress || contact.displayName"
            class="contact-row"
            :class="{ active: selectedContact?.smtpAddress === contact.smtpAddress }"
            type="button"
            @click="selectedContact = contact"
          >
            <strong>{{ contactTitle(contact) }}</strong>
            <span>{{ contact.smtpAddress || '-' }}</span>
            <small>{{ contact.mailCount }} mails / {{ contact.calendarCount }} calendar</small>
          </button>

          <el-empty v-if="!loading && contacts.length === 0" description="目前沒有符合條件的聯絡人" />
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
                <el-button :icon="Search" :loading="loading" @click="lookupContact" />
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
