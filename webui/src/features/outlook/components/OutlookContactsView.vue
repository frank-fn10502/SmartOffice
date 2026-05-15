<script setup lang="ts">
import { computed, onMounted, ref, watch } from 'vue'
import { ArrowRight, Refresh, Search, UserFilled } from '@element-plus/icons-vue'
import { ElMessage } from 'element-plus'
import { normalizeAddressBookContact, outlookApi } from '../api/outlook'
import type { AddressBookContactDto, AddressBookRootDto } from '../models/outlook'
import { fetchResultEndpoint, requestIdFromResponse } from '../composables/outlookRequests'
import type { OutlookDashboardState } from '../composables/useOutlookDashboard'

const props = defineProps<{
  dashboard?: OutlookDashboardState
}>()

type ContactNode = {
  contact: AddressBookContactDto
  depth: number
}

const contacts = ref<AddressBookContactDto[]>([])
const roots = ref<AddressBookRootDto[]>([])
const selectedRoot = ref<AddressBookRootDto | null>(null)
const selectedContact = ref<AddressBookContactDto | null>(null)
const expandedGroups = ref<Record<string, AddressBookContactDto[]>>({})
const loadingGroups = ref<Record<string, boolean>>({})
const query = ref('')
const loading = ref(false)
const searching = ref(false)
const message = ref('尚未載入通訊錄。')
const nextOffset = ref(0)
const totalCount = ref(0)
const hasMore = ref(false)
let initialLoadRequested = false
let searchTimer: number | undefined
let searchRunId = 0
const pollDelayMs = 500
const pollTimeoutMs = 120000

const peopleCount = computed(() => contacts.value.filter((contact) => !contact.isGroup).length)
const groupCount = computed(() => contacts.value.filter((contact) => contact.isGroup).length)
const visibleNodes = computed(() => flattenContacts(sortContacts(filterContacts(contacts.value))))
const selectedTitle = computed(() => selectedContact.value ? contactTitle(selectedContact.value) : '尚未選取')

function contactKey(contact: AddressBookContactDto) {
  return (contact.smtpAddress || contact.rawAddress || contact.id || contact.displayName).trim().toLowerCase()
}

function contactTitle(contact: AddressBookContactDto) {
  return contact.displayName || contact.smtpAddress || contact.rawAddress || '(未命名)'
}

function rootKey(root: AddressBookRootDto) {
  return (root.id || root.name || root.source).trim().toLowerCase()
}

function rootTitle(root: AddressBookRootDto) {
  return root.name || root.id || '(來源)'
}

function contactTypeLabel(contact: AddressBookContactDto) {
  return contact.isGroup ? '群組' : '人員'
}

function relationLabel(contact: AddressBookContactDto) {
  if (contact.isLikelySelf) return '你'
  if (contact.isGroup && contact.isRelatedToSelf) return '你所在的群組'
  if (!contact.isGroup && contact.isRelatedToSelf) return '與你同組'
  return ''
}

function groupMembers(contact: AddressBookContactDto) {
  return expandedGroups.value[contactKey(contact)] ?? []
}

function isExpanded(contact: AddressBookContactDto) {
  return groupMembers(contact).length > 0
}

function isLoadingGroup(contact: AddressBookContactDto) {
  return Boolean(loadingGroups.value[contactKey(contact)])
}

function isSelected(contact: AddressBookContactDto) {
  return Boolean(selectedContact.value && contactKey(selectedContact.value) === contactKey(contact))
}

function sortScore(contact: AddressBookContactDto) {
  return (contact.isLikelySelf ? 100000 : 0)
    + (contact.isRelatedToSelf ? 50000 : 0)
    + (contact.isGroup ? 5000 : 0)
    + contact.relationScore * 10
    + contact.mailCount
    + contact.calendarCount
}

function sortContacts(source: AddressBookContactDto[]) {
  return [...source].sort((a, b) => sortScore(b) - sortScore(a)
    || contactTitle(a).localeCompare(contactTitle(b), 'zh-Hant'))
}

function filterContacts(source: AddressBookContactDto[]) {
  const text = query.value.trim().toLowerCase()
  if (!text) return source
  return source.filter((contact) =>
    contactTitle(contact).toLowerCase().includes(text)
    || contact.smtpAddress.toLowerCase().includes(text)
    || contact.domain.toLowerCase().includes(text))
}

function flattenContacts(source: AddressBookContactDto[]) {
  const nodes: ContactNode[] = []
  const visited = new Set<string>()
  for (const contact of source) addNode(nodes, contact, 0, visited)
  return nodes
}

function addNode(nodes: ContactNode[], contact: AddressBookContactDto, depth: number, visited: Set<string>) {
  const key = contactKey(contact)
  if (!key || visited.has(`${depth}:${key}`)) return
  nodes.push({ contact, depth })
  if (!contact.isGroup || !isExpanded(contact)) return
  visited.add(`${depth}:${key}`)
  for (const member of sortContacts(groupMembers(contact))) addNode(nodes, member, depth + 1, visited)
}

function mergeContacts(nextContacts: AddressBookContactDto[]) {
  const byKey = new Map<string, AddressBookContactDto>()
  for (const contact of contacts.value) byKey.set(contactKey(contact), contact)
  for (const contact of nextContacts) {
    const key = contactKey(contact)
    if (!key) continue
    byKey.set(key, { ...byKey.get(key), ...contact })
  }
  contacts.value = Array.from(byKey.values())
}

function selectContact(contact: AddressBookContactDto) {
  selectedContact.value = contact
  mergeContacts([contact])
}

function normalizeRoots(loadedRoots: AddressBookRootDto[]) {
  const seen = new Map<string, AddressBookRootDto>()
  for (const root of loadedRoots) {
    const key = `${rootTitle(root)}|${root.addressListType}|${root.entryCount}`.toLowerCase()
    if (!seen.has(key)) seen.set(key, root)
  }
  return Array.from(seen.values())
}

async function loadContacts() {
  if (loading.value) return
  loading.value = true
  message.value = '正在整理通訊錄...'
  try {
    contacts.value = []
    expandedGroups.value = {}
    selectedContact.value = null
    await loadKnownContacts()
    await loadRoots()
    if (roots.value[0]) await loadRootPage(roots.value[0], 0)
    await loadKnownContacts()
    if (!selectedContact.value) selectedContact.value = visibleNodes.value[0]?.contact ?? null
  } catch (error) {
    message.value = error instanceof Error ? error.message : '通訊錄載入失敗。'
  } finally {
    loading.value = false
  }
}

async function loadKnownContacts() {
  const page = await outlookApi.getAddressBookContacts('', 5000)
  mergeContacts(page.contacts)
  message.value = `已載入 ${contacts.value.length} 筆聯絡人。`
}

async function loadRoots() {
  const response = await outlookApi.requestAddressBookRoots()
  const requestId = requestIdFromResponse(response)
  if (!requestId) return
  const state = await pollFetchResult<{ roots?: unknown[] }>(fetchResultEndpoint(response), requestId, 500, '通訊錄來源載入失敗。')
  roots.value = normalizeRoots(outlookApi.normalizeAddressBookRootsData(state.data).roots)
}

async function loadRootPage(root: AddressBookRootDto, offset = 0) {
  selectedRoot.value = root
  message.value = `正在載入 ${rootTitle(root)}...`
  const response = await outlookApi.requestAddressListEntries({
    addressListId: root.id,
    addressListName: root.name,
    offset,
    pageSize: 100,
  })
  const requestId = requestIdFromResponse(response)
  if (!requestId) return
  const state = await pollFetchResult<Record<string, unknown>>(fetchResultEndpoint(response), requestId, 500, '通訊錄載入失敗。')
  const page = outlookApi.normalizeAddressListEntriesData(state.data)
  mergeContacts(page.contacts)
  nextOffset.value = page.offset + page.contacts.length
  totalCount.value = page.totalCount
  hasMore.value = page.hasMore
  message.value = `${rootTitle(root)} ${nextOffset.value}/${page.totalCount} 筆。`
}

async function toggleGroup(contact: AddressBookContactDto) {
  if (!contact.isGroup || isLoadingGroup(contact)) return
  const key = contactKey(contact)
  if (isExpanded(contact)) {
    const next = { ...expandedGroups.value }
    delete next[key]
    expandedGroups.value = next
    return
  }
  loadingGroups.value = { ...loadingGroups.value, [key]: true }
  try {
    const response = await outlookApi.requestAddressBookGroupMembers({
      groupId: contact.id,
      groupSmtpAddress: contact.smtpAddress,
      maxMembers: 5000,
    })
    const members = response.state === 'completed' && response.data?.members
      ? response.data.members
      : await streamGroupMembers(response)
    expandedGroups.value = { ...expandedGroups.value, [key]: members }
    mergeContacts(members)
  } catch (error) {
    ElMessage.error(error instanceof Error ? error.message : '群組展開失敗。')
  } finally {
    const next = { ...loadingGroups.value }
    delete next[key]
    loadingGroups.value = next
  }
}

async function streamGroupMembers(response: { requestId?: string; request?: string; data?: unknown }) {
  const requestId = requestIdFromResponse(response)
  if (!requestId) return []
  const started = Date.now()
  const seen = new Map<string, AddressBookContactDto>()
  let cursor = ''
  while (Date.now() - started < pollTimeoutMs) {
    const state = await outlookApi.fetchResult<{ members?: unknown[] }>(fetchResultEndpoint(response), { requestId, cursor, take: 100 })
    const members = Array.isArray(state.data?.members) ? state.data.members.map(normalizeAddressBookContact) : []
    for (const member of members) seen.set(contactKey(member), member)
    if (state.next.hasMore) {
      cursor = state.next.cursor
      continue
    }
    if (state.state === 'completed') return Array.from(seen.values())
    if (state.state && !['accepted', 'running'].includes(state.state)) throw new Error(state.message || '群組展開失敗。')
    await delay(pollDelayMs)
  }
  throw new Error('群組展開逾時。')
}

async function pollFetchResult<TData>(endpoint: string, requestId: string, take: number, failureMessage: string) {
  const started = Date.now()
  while (Date.now() - started < pollTimeoutMs) {
    const state = await outlookApi.fetchResult<TData>(endpoint, { requestId, cursor: '', take })
    if (state.state === 'completed') return state
    if (state.state && !['accepted', 'running'].includes(state.state)) throw new Error(state.message || failureMessage)
    await delay(pollDelayMs)
  }
  throw new Error('通訊錄載入逾時。')
}

function relationContacts(data: unknown) {
  const source = (data ?? {}) as Record<string, unknown>
  const keys = ['target', 'matches', 'members', 'memberGroups', 'memberOfGroups', 'containingGroups']
  const result: AddressBookContactDto[] = []
  for (const key of keys) {
    const value = source[key] ?? source[key[0].toUpperCase() + key.slice(1)]
    if (Array.isArray(value)) result.push(...value.map(normalizeAddressBookContact))
    else if (value) result.push(normalizeAddressBookContact(value))
  }
  return result
}

async function searchContacts(text: string, runId: number) {
  searching.value = true
  try {
    const response = await outlookApi.requestAddressBookRelation({ query: text, take: 100 })
    const requestId = requestIdFromResponse(response)
    if (!requestId) return
    const state = await pollFetchResult<Record<string, unknown>>(fetchResultEndpoint(response), requestId, 100, '通訊錄搜尋失敗。')
    if (runId !== searchRunId) return
    mergeContacts(relationContacts(state.data))
  } finally {
    if (runId === searchRunId) searching.value = false
  }
}

function scheduleSearch() {
  window.clearTimeout(searchTimer)
  const text = query.value.trim()
  if (text.length < 2) {
    searchRunId += 1
    searching.value = false
    return
  }
  const runId = ++searchRunId
  searchTimer = window.setTimeout(() => {
    void searchContacts(text, runId).catch((error) => {
      if (runId === searchRunId) message.value = error instanceof Error ? error.message : '通訊錄搜尋失敗。'
    })
  }, 300)
}

function requestInitialLoad(reason: 'active-view' | 'background') {
  if (initialLoadRequested || loading.value || contacts.value.length > 0) return
  if (reason === 'background' && !props.dashboard?.outlookFirstLoadCompleted.value) return
  if (reason === 'active-view' && props.dashboard?.activeView.value !== 'contacts') return
  initialLoadRequested = true
  void loadContacts().catch(() => { initialLoadRequested = false })
}

function delay(ms: number) {
  return new Promise((resolve) => window.setTimeout(resolve, ms))
}

onMounted(() => {
  if (!props.dashboard) void loadContacts()
  else requestInitialLoad('active-view')
})

watch(() => props.dashboard?.activeView.value, () => requestInitialLoad('active-view'))
watch(
  () => props.dashboard?.outlookFirstLoadCompleted.value,
  (completed) => {
    if (completed) window.setTimeout(() => requestInitialLoad('background'), 1200)
  },
  { immediate: true },
)
watch(query, scheduleSearch)
</script>

<template>
  <main class="contacts-layout">
    <section class="panel contacts-panel">
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
            :prefix-icon="Search"
            placeholder="搜尋姓名或 email"
          />
          <el-button :icon="Refresh" :loading="loading" :disabled="loading" @click="loadContacts">重新整理</el-button>
        </div>
      </div>

      <div class="contacts-status" :class="{ loading: loading || searching }">
        <strong>{{ loading || searching ? '通訊錄整理中' : '通訊錄狀態' }}</strong>
        <span>{{ message }}</span>
        <span>{{ peopleCount }} 人員</span>
        <span>{{ groupCount }} 群組</span>
      </div>

      <div class="contacts-page">
        <nav class="contacts-roots" aria-label="通訊錄來源">
          <strong>來源</strong>
          <button
            v-for="root in roots"
            :key="rootKey(root)"
            class="root-row"
            :class="{ active: selectedRoot && rootKey(selectedRoot) === rootKey(root) }"
            type="button"
            @click="loadRootPage(root, 0)"
          >
            <span>{{ rootTitle(root) }}</span>
            <small>{{ root.entryCount }} 筆</small>
          </button>
          <span v-if="!roots.length" class="contacts-empty">尚未載入來源</span>
        </nav>

        <section class="contacts-list" aria-label="聯絡人">
          <button
            v-if="selectedRoot && hasMore"
            class="load-next-page"
            type="button"
            :disabled="loading"
            @click="loadRootPage(selectedRoot, nextOffset)"
          >
            載入下一頁
          </button>

          <div v-if="loading || searching" class="contacts-loading">
            {{ searching ? '正在搜尋通訊錄...' : '正在載入通訊錄...' }}
          </div>

          <div v-for="{ contact, depth } in visibleNodes" :key="`${depth}:${contactKey(contact)}`" class="contact-line" :style="{ '--depth': depth }">
            <button
              class="contact-expander"
              :class="{ expanded: isExpanded(contact), hidden: !contact.isGroup }"
              type="button"
              :disabled="!contact.isGroup || isLoadingGroup(contact)"
              @click="toggleGroup(contact)"
            >
              <el-icon v-if="contact.isGroup"><ArrowRight /></el-icon>
            </button>
            <button
              class="contact-row"
              :class="{ active: isSelected(contact), group: contact.isGroup }"
              type="button"
              @click="selectContact(contact)"
            >
              <span class="contact-main">
                <strong>{{ contactTitle(contact) }}</strong>
                <small>{{ contact.smtpAddress || '-' }}</small>
              </span>
              <span class="contact-tags">
                <el-tag size="small" effect="plain">{{ contactTypeLabel(contact) }}</el-tag>
                <el-tag v-if="relationLabel(contact)" size="small" type="success" effect="plain">{{ relationLabel(contact) }}</el-tag>
              </span>
            </button>
          </div>

          <el-empty v-if="!loading && visibleNodes.length === 0" description="目前沒有符合條件的聯絡人" />
        </section>

        <aside class="contact-detail">
          <header>
            <strong>{{ selectedTitle }}</strong>
            <el-tag v-if="selectedContact" effect="plain">{{ contactTypeLabel(selectedContact) }}</el-tag>
          </header>
          <template v-if="selectedContact">
            <dl>
              <dt>Email</dt>
              <dd>{{ selectedContact.smtpAddress || '-' }}</dd>
              <dt>關係</dt>
              <dd>{{ relationLabel(selectedContact) || '一般聯絡人' }}</dd>
              <dt>公司 / 部門</dt>
              <dd>{{ selectedContact.companyName || '-' }} / {{ selectedContact.department || '-' }}</dd>
              <dt>職稱</dt>
              <dd>{{ selectedContact.jobTitle || '-' }}</dd>
              <dt>互動</dt>
              <dd>{{ selectedContact.mailCount }} mails / {{ selectedContact.calendarCount }} calendar</dd>
              <dt v-if="selectedContact.isGroup">成員</dt>
              <dd v-if="selectedContact.isGroup">{{ selectedContact.memberCount || selectedContact.memberSmtpAddresses.length }} 位</dd>
            </dl>
          </template>
          <span v-else class="contacts-empty">選取聯絡人查看詳細資料。</span>
        </aside>
      </div>
    </section>
  </main>
</template>
