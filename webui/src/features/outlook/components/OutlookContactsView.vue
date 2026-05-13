<script setup lang="ts">
import { computed, onMounted, ref } from 'vue'
import { Refresh, Search, UserFilled } from '@element-plus/icons-vue'
import { ElMessage } from 'element-plus'
import { outlookApi } from '../api/outlook'
import type { AddressBookContactDto } from '../models/outlook'
import { collectOutlookRequestData, waitForOutlookRequest } from '../composables/outlookRequests'
import { formatDateTime } from '../utils/formatters'

const contacts = ref<AddressBookContactDto[]>([])
const selectedContact = ref<AddressBookContactDto | null>(null)
const query = ref('')
const loading = ref(false)
const lookupEmail = ref('')
const lookupMessage = ref('')

const knownCount = computed(() => contacts.value.filter((contact) => contact.isKnown).length)
const selfCount = computed(() => contacts.value.filter((contact) => contact.isLikelySelf).length)

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
  loading.value = true
  try {
    const response = await outlookApi.requestAddressBook({
      includeOutlookContacts: true,
      includeAddressLists: true,
      maxContacts: 1000,
      maxAddressEntriesPerList: 500,
    })
    await waitForOutlookRequest(response, { timeoutMs: 120000 })
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
  } finally {
    loading.value = false
  }
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
  loading.value = true
  try {
    await loadContacts()
    ElMessage.success('通訊錄同步完成')
  } finally {
    loading.value = false
  }
}

onMounted(() => {
  void loadContacts()
})
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
            :prefix-icon="Search"
            placeholder="姓名、email 或 domain"
            @keyup.enter="loadContacts"
            @clear="loadContacts"
          />
          <el-button :icon="Refresh" :loading="loading" @click="loadContacts">重新整理</el-button>
          <el-button type="primary" :loading="loading" @click="syncAddressBook">同步 Outlook 通訊錄</el-button>
        </div>
      </div>

      <div class="contacts-page">
        <div class="contacts-list">
          <div class="contacts-stats">
            <span>已知關聯 {{ knownCount }}</span>
            <span>可能是自己 {{ selfCount }}</span>
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
              <el-tag v-else type="success" effect="plain">已知關聯</el-tag>
            </div>

            <div class="rule-detail">
              <span>Email：{{ selectedContact.smtpAddress || '-' }}</span>
              <span>Domain：{{ selectedContact.domain || '-' }}</span>
              <span>來源：{{ selectedContact.sources.join(', ') || selectedContact.source || '-' }}</span>
              <span>公司 / 部門：{{ selectedContact.companyName || '-' }} / {{ selectedContact.department || '-' }}</span>
              <span>職稱：{{ selectedContact.jobTitle || '-' }}</span>
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
