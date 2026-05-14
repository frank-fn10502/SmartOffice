<script setup lang="ts">
import { computed, onMounted, ref, watch } from 'vue'
import { ArrowDown, Refresh, Search, UserFilled } from '@element-plus/icons-vue'
import { ElMessage } from 'element-plus'
import { normalizeAddressBookContact, outlookApi } from '../api/outlook'
import type { AddressBookContactDto, AddressBookRootDto } from '../models/outlook'
import { fetchResultEndpoint, requestIdFromResponse } from '../composables/outlookRequests'
import { formatDateTime } from '../utils/formatters'
import type { OutlookDashboardState } from '../composables/useOutlookDashboard'

const props = defineProps<{
  dashboard?: OutlookDashboardState
}>()

const contacts = ref<AddressBookContactDto[]>([])
const addressBookRoots = ref<AddressBookRootDto[]>([])
const selectedRoot = ref<AddressBookRootDto | null>(null)
const addressListOffset = ref(0)
const addressListHasMore = ref(false)
const addressListTotal = ref(0)
const selectedContact = ref<AddressBookContactDto | null>(null)
const query = ref('')
const loadingContacts = ref(false)
const lookupLoading = ref(false)
const lastUpdatedAt = ref<Date | null>(null)
const loadMessage = ref('尚未載入通訊錄。')
const lookupEmail = ref('')
const lookupMessage = ref('')
const groupMembersByKey = ref<Record<string, AddressBookContactDto[]>>({})
const expandingGroups = ref<Record<string, boolean>>({})
let initialLoadRequested = false
const pollDelayMs = 500
const pollTimeoutMs = 120000

const knownCount = computed(() => contacts.value.filter((contact) => contact.isKnown).length)
const selfCount = computed(() => contacts.value.filter((contact) => contact.isLikelySelf).length)
const groupCount = computed(() => contacts.value.filter((contact) => contact.isGroup).length)
const personCount = computed(() => contacts.value.filter((contact) => !contact.isGroup).length)
const selectedRootKey = computed(() => selectedRoot.value ? rootKey(selectedRoot.value) : '')
const visibleContacts = computed(() => filterContacts(contacts.value))
const lastUpdatedText = computed(() => lastUpdatedAt.value
  ? lastUpdatedAt.value.toLocaleString('zh-TW', { hour12: false })
  : '尚未完成同步')
const selectedGroupMembers = computed(() => selectedContact.value
  ? groupMembersByKey.value[contactKey(selectedContact.value)] ?? selectedContact.value.memberSmtpAddresses.map(memberEmailToContact)
  : [])
const selectedGroupExpanded = computed(() => Boolean(selectedContact.value && (selectedContact.value.groupMembersLoaded || groupMembersByKey.value[contactKey(selectedContact.value)]?.length)))
const selectedGroupLoading = computed(() => Boolean(selectedContact.value && (selectedContact.value.groupMembersLoading || expandingGroups.value[contactKey(selectedContact.value)])))

function contactTitle(contact: AddressBookContactDto) {
  return contact.displayName || contact.smtpAddress || '(unknown)'
}

function rootTitle(root: AddressBookRootDto) {
  return root.name || root.id || '(address list)'
}

function rootKey(root: AddressBookRootDto) {
  return (root.id || root.name || root.source).trim().toLowerCase()
}

function contactKey(contact: AddressBookContactDto) {
  return (contact.smtpAddress || contact.rawAddress || contact.id || contact.displayName).trim().toLowerCase()
}

function isSelectedRoot(root: AddressBookRootDto) {
  return selectedRootKey.value === rootKey(root)
}

function isSelectedContact(contact: AddressBookContactDto) {
  return Boolean(selectedContact.value && contactKey(selectedContact.value) === contactKey(contact))
}

function memberEmailToContact(email: string): AddressBookContactDto {
  return emptyContact(email, selectedContact.value?.memberGroupSmtpAddresses.includes(email) ?? false)
}

function groupMembersFor(contact: AddressBookContactDto) {
  return groupMembersByKey.value[contactKey(contact)] ?? contact.memberSmtpAddresses.map((email) => emptyContact(email, contact.memberGroupSmtpAddresses.includes(email)))
}

function groupExpanded(contact: AddressBookContactDto) {
  return Boolean(contact.groupMembersLoaded || groupMembersByKey.value[contactKey(contact)]?.length)
}

function groupLoading(contact: AddressBookContactDto) {
  return Boolean(contact.groupMembersLoading || expandingGroups.value[contactKey(contact)])
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

function emptyContact(email: string, isGroup = false): AddressBookContactDto {
  return {
    id: email,
    displayName: email,
    smtpAddress: email,
    rawAddress: email,
    addressType: 'SMTP',
    entryUserType: isGroup ? 'olExchangeDistributionListAddressEntry' : '',
    source: 'group_member',
    companyName: '',
    jobTitle: '',
    department: '',
    officeLocation: '',
    businessTelephoneNumber: '',
    mobileTelephoneNumber: '',
    domain: email.includes('@') ? email.split('@').pop() ?? '' : '',
    isKnown: true,
    isLikelySelf: false,
    isRelatedToSelf: false,
    isGroup,
    memberCount: 0,
    groupMembersLoaded: false,
    groupMembersLoading: false,
    groupMembersRequestId: '',
    relationScore: 0,
    mailCount: 0,
    calendarCount: 0,
    senderCount: 0,
    recipientCount: 0,
    organizerCount: 0,
    attendeeCount: 0,
    groupMemberCount: 0,
    relationKinds: ['group_member'],
    sources: ['group_member'],
    memberSmtpAddresses: [],
    memberGroupSmtpAddresses: [],
    memberOfGroupSmtpAddresses: selectedContact.value?.smtpAddress ? [selectedContact.value.smtpAddress] : [],
    folderPaths: [],
    recentMailIds: [],
    sampleSubjects: [],
  }
}

async function loadContacts() {
  if (loadingContacts.value) return
  loadingContacts.value = true
  loadMessage.value = '正在載入 Outlook 通訊錄來源...'
  try {
    const response = await outlookApi.requestAddressBookRoots()
    contacts.value = []
    addressBookRoots.value = []
    selectedRoot.value = null
    addressListOffset.value = 0
    addressListHasMore.value = false
    addressListTotal.value = 0
    selectedContact.value = null
    await streamAddressBookRootsFromRequest(response)
    const firstRoot = addressBookRoots.value[0]
    if (firstRoot) {
      await loadAddressListEntriesPage(firstRoot, 0)
    }
    lastUpdatedAt.value = new Date()
    if (!firstRoot) loadMessage.value = `通訊錄來源載入完成，共 ${addressBookRoots.value.length} 個來源。`
  } catch (error) {
    loadMessage.value = error instanceof Error ? error.message : '通訊錄同步失敗。'
    throw error
  } finally {
    loadingContacts.value = false
  }
}

async function streamAddressBookRootsFromRequest(response: { requestId?: string; request?: string; data?: unknown }) {
  const requestId = requestIdFromResponse(response)
  if (!requestId) return
  const endpoint = fetchResultEndpoint(response)
  const state = await pollFetchResult<{ roots?: unknown[] }>(endpoint, requestId, '', 500, '通訊錄來源載入失敗。')
  addressBookRoots.value = outlookApi.normalizeAddressBookRootsData(state.data).roots
}

async function loadAddressListEntries(root: AddressBookRootDto, offset = 0) {
  if (loadingContacts.value) return
  loadingContacts.value = true
  try {
    await loadAddressListEntriesPage(root, offset)
  } finally {
    loadingContacts.value = false
  }
}

async function loadAddressListEntriesPage(root: AddressBookRootDto, offset = 0) {
  loadMessage.value = `正在載入 ${rootTitle(root)} entries...`
  const response = await outlookApi.requestAddressListEntries({
    addressListId: root.id,
    addressListName: root.name,
    offset,
    pageSize: 100,
  })
  const requestId = requestIdFromResponse(response)
  if (!requestId) return
  const state = await pollFetchResult<Record<string, unknown>>(fetchResultEndpoint(response), requestId, '', 500, '通訊錄 entries 載入失敗。')
  const page = outlookApi.normalizeAddressListEntriesData(state.data)
  selectedRoot.value = root
  addressListOffset.value = page.offset + page.contacts.length
  addressListHasMore.value = page.hasMore
  addressListTotal.value = page.totalCount
  contacts.value = offset > 0 ? [...contacts.value, ...page.contacts] : page.contacts
  selectedContact.value = selectedContact.value && contacts.value.some((contact) => contactKey(contact) === contactKey(selectedContact.value!))
    ? selectedContact.value
    : contacts.value[0] ?? null
  loadMessage.value = `${rootTitle(root)} 已載入 ${contacts.value.length}/${page.totalCount} 筆。`
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

async function pollFetchResult<TData>(
  endpoint: string,
  requestId: string,
  cursor: string,
  take: number,
  failureMessage: string,
) {
  const started = Date.now()
  while (Date.now() - started < pollTimeoutMs) {
    const state = await outlookApi.fetchResult<TData>(endpoint, { requestId, cursor, take })
    if (state.state === 'completed') return state
    if (state.state && !['accepted', 'running'].includes(state.state)) throw new Error(state.message || failureMessage)
    await new Promise((resolve) => window.setTimeout(resolve, pollDelayMs))
  }

  throw new Error('通訊錄載入逾時。')
}

function mergeContacts(target: Map<string, AddressBookContactDto>, nextContacts: AddressBookContactDto[]) {
  for (const contact of nextContacts) {
    const key = contact.smtpAddress || contact.rawAddress || contact.displayName || contact.id
    if (key) target.set(key.toLowerCase(), contact)
  }
}

function upsertContacts(nextContacts: AddressBookContactDto[]) {
  const seen = new Map<string, AddressBookContactDto>()
  mergeContacts(seen, contacts.value)
  mergeContacts(seen, nextContacts)
  contacts.value = Array.from(seen.values())
}

function updateSelectedGroupMembers(group: AddressBookContactDto, members: AddressBookContactDto[]) {
  const key = contactKey(group)
  groupMembersByKey.value = { ...groupMembersByKey.value, [key]: members }
  upsertContacts(members)
  const updated = {
    ...group,
    groupMembersLoaded: true,
    groupMembersLoading: false,
    memberCount: Math.max(group.memberCount, members.length),
    memberSmtpAddresses: members.map((member) => member.smtpAddress).filter(Boolean).slice(0, 50),
    memberGroupSmtpAddresses: members.filter((member) => member.isGroup).map((member) => member.smtpAddress).filter(Boolean).slice(0, 50),
  }
  selectedContact.value = updated
  upsertContacts([updated])
}

async function expandSelectedGroup(forceRefresh = false) {
  const group = selectedContact.value
  if (!group?.isGroup || selectedGroupLoading.value) return
  const key = contactKey(group)
  expandingGroups.value = { ...expandingGroups.value, [key]: true }
  selectedContact.value = { ...group, groupMembersLoading: true }
  try {
    const response = await outlookApi.requestAddressBookGroupMembers({
      groupId: group.id,
      groupSmtpAddress: group.smtpAddress,
      maxMembers: 5000,
      forceRefresh,
    })
    if (response.state === 'completed' && response.data?.members) {
      updateSelectedGroupMembers(group, response.data.members)
      ElMessage.success('群組成員載入完成')
      return
    }

    const members = await streamGroupMembersFromRequest(response)
    updateSelectedGroupMembers(group, members)
    ElMessage.success('群組成員載入完成')
  } finally {
    const next = { ...expandingGroups.value }
    delete next[key]
    expandingGroups.value = next
    if (selectedContact.value && contactKey(selectedContact.value) === key) {
      selectedContact.value = { ...selectedContact.value, groupMembersLoading: false }
    }
  }
}

async function expandGroupFromList(contact: AddressBookContactDto, forceRefresh = false) {
  selectedContact.value = contact
  await expandSelectedGroup(forceRefresh)
}

async function streamGroupMembersFromRequest(response: { requestId?: string; request?: string; data?: unknown }) {
  const requestId = requestIdFromResponse(response)
  if (!requestId) return []
  const endpoint = fetchResultEndpoint(response)
  const started = Date.now()
  const timeoutMs = 120000
  const seen = new Map<string, AddressBookContactDto>()
  let cursor = ''

  while (Date.now() - started < timeoutMs) {
    const state = await outlookApi.fetchResult<{ members?: unknown[] }>(endpoint, { requestId, cursor, take: 100 })
    const rawMembers = Array.isArray(state.data?.members) ? state.data.members : []
    mergeContacts(seen, rawMembers.map(normalizeAddressBookContact))
    if (state.next.hasMore) {
      cursor = state.next.cursor
      continue
    }
    if (state.state === 'completed') return Array.from(seen.values())
    if (state.state && !['accepted', 'running'].includes(state.state)) throw new Error(state.message || '群組成員載入失敗。')
    await new Promise((resolve) => window.setTimeout(resolve, 500))
  }

  throw new Error('群組成員載入逾時。')
}

function selectContact(contact: AddressBookContactDto) {
  selectedContact.value = contact
  if (!contacts.value.some((item) => contactKey(item) === contactKey(contact))) {
    contacts.value = [contact, ...contacts.value]
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
          />
          <el-button :icon="Refresh" :loading="loadingContacts" :disabled="loadingContacts" @click="loadContacts">重新整理</el-button>
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
        <nav class="contacts-source-tree" aria-label="通訊錄來源">
          <div class="contacts-tree-title">
            <strong>來源</strong>
            <span>{{ addressBookRoots.length }} 個 AddressList</span>
          </div>
          <button
            v-for="root in addressBookRoots"
            :key="root.id || root.name"
            class="source-row"
            :class="{ active: isSelectedRoot(root) }"
            type="button"
            @click="loadAddressListEntries(root)"
          >
            <span class="source-branch" />
            <span class="source-main">
              <strong>{{ rootTitle(root) }}</strong>
              <small>{{ root.addressListType || root.source || '-' }}</small>
            </span>
            <el-tag size="small" effect="plain">{{ root.entryCount }}</el-tag>
          </button>
          <el-empty v-if="!loadingContacts && addressBookRoots.length === 0" description="尚未載入來源" />
        </nav>

        <div class="contacts-list">
          <div class="contacts-stats">
            <span>{{ selectedRoot ? rootTitle(selectedRoot) : '尚未選擇來源' }}</span>
            <span v-if="selectedRoot">{{ contacts.length }}/{{ addressListTotal || contacts.length }} entries</span>
            <span>個人 {{ personCount }}</span>
            <span>Group {{ groupCount }}</span>
            <span>已知關聯 {{ knownCount }}</span>
            <span>可能是自己 {{ selfCount }}</span>
          </div>

          <el-button
            v-if="selectedRoot && addressListHasMore"
            class="load-next-page"
            :loading="loadingContacts"
            :disabled="loadingContacts"
            @click="loadAddressListEntries(selectedRoot, addressListOffset)"
          >
            載入下一頁
          </el-button>

          <div v-if="loadingContacts" class="contacts-loading" role="status">
            <span />
            <strong>正在載入通訊錄...</strong>
          </div>

          <template v-for="contact in visibleContacts" :key="contactKey(contact)">
            <div class="contact-tree-node">
              <span class="tree-connector" />
              <button
                class="contact-row"
                :class="{ active: isSelectedContact(contact), group: contact.isGroup }"
                type="button"
                @click="selectContact(contact)"
              >
                <span class="contact-row-main">
                  <strong>{{ contactTitle(contact) }}</strong>
                  <span>{{ contact.smtpAddress || '-' }}</span>
                </span>
                <span class="contact-row-meta">
                  <el-tag v-if="contact.isGroup" size="small" effect="plain">Group</el-tag>
                  <el-tag v-if="contact.isGroup && contact.isRelatedToSelf" size="small" type="success" effect="plain">包含自己</el-tag>
                  <small>{{ contact.mailCount }} mails / {{ contact.calendarCount }} calendar</small>
                </span>
              </button>
              <el-button
                v-if="contact.isGroup"
                class="group-row-action"
                :icon="ArrowDown"
                :loading="groupLoading(contact)"
                :disabled="groupLoading(contact)"
                @click.stop="expandGroupFromList(contact)"
              >
                {{ groupExpanded(contact) ? '已展開' : '展開' }}
              </el-button>
            </div>

            <div
              v-if="contact.isGroup && groupExpanded(contact)"
              class="contact-children"
            >
              <button
                v-for="member in groupMembersFor(contact)"
                :key="`${contactKey(contact)}:${contactKey(member)}`"
                class="contact-child-row"
                type="button"
                @click="selectContact(member)"
              >
                <span class="tree-connector child" />
                <strong>{{ contactTitle(member) }}</strong>
                <span>{{ member.smtpAddress || '-' }}</span>
                <el-tag v-if="member.isGroup" size="small" effect="plain">子群組</el-tag>
              </button>
            </div>
          </template>

          <el-empty
            v-if="!loadingContacts && visibleContacts.length === 0"
            :description="selectedRoot ? '目前沒有符合條件的聯絡人' : '請先選擇左側通訊錄來源'"
          />
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
              <span v-if="selectedContact.isGroup">
                與自己關聯：{{ selectedContact.isRelatedToSelf ? '包含自己' : '尚未確認' }}
              </span>
              <span v-if="selectedContact.isGroup">成員數：{{ selectedContact.memberCount }}</span>
              <span v-if="selectedContact.isGroup">
                展開狀態：{{ selectedGroupExpanded ? '已載入' : selectedGroupLoading ? '載入中' : '未展開' }}
              </span>
              <span v-if="selectedContact.memberOfGroupSmtpAddresses.length > 0">
                隸屬群組：{{ selectedContact.memberOfGroupSmtpAddresses.slice(0, 5).join(', ') }}
              </span>
              <span>最近互動：{{ selectedContact.lastSeen ? formatDateTime(selectedContact.lastSeen) : '-' }}</span>
              <span>最早出現：{{ selectedContact.firstSeen ? formatDateTime(selectedContact.firstSeen) : '-' }}</span>
              <span>關聯分數：{{ selectedContact.relationScore }}</span>
            </div>

            <div v-if="selectedContact.isGroup" class="group-members-panel">
              <div class="group-members-header">
                <strong>Group members</strong>
                <div>
                  <el-button :icon="ArrowDown" :loading="selectedGroupLoading" :disabled="selectedGroupLoading" @click="expandSelectedGroup(false)">
                    {{ selectedGroupExpanded ? '已展開' : '展開成員' }}
                  </el-button>
                  <el-button :loading="selectedGroupLoading" :disabled="selectedGroupLoading" @click="expandSelectedGroup(true)">重新展開</el-button>
                </div>
              </div>
              <div v-if="selectedGroupMembers.length > 0" class="group-members-list">
                <button
                  v-for="member in selectedGroupMembers"
                  :key="contactKey(member)"
                  class="group-member-row"
                  type="button"
                  @click="selectContact(member)"
                >
                  <strong>{{ contactTitle(member) }}</strong>
                  <span>{{ member.smtpAddress || '-' }}</span>
                  <el-tag v-if="member.isGroup" size="small" effect="plain">子群組</el-tag>
                </button>
              </div>
              <span v-else class="group-members-empty">尚未展開；點「展開成員」後會載入 direct members。</span>
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
