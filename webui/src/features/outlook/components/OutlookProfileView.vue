<script setup lang="ts">
import { computed, onMounted, ref, watch } from 'vue'
import { Refresh, UserFilled } from '@element-plus/icons-vue'
import { outlookApi } from '../api/outlook'
import type { OutlookDashboardState } from '../composables/useOutlookDashboard'
import type { AddressBookContactDto, OutlookProfileDto, OutlookStoreDto } from '../models/outlook'

const props = defineProps<{
  dashboard: OutlookDashboardState
}>()

const profile = ref<OutlookProfileDto | null>(null)
const loadingProfile = ref(false)
const profileMessage = ref('等待 Outlook 郵件載入完成。')

const mailStats = computed(() => profile.value?.mailStats ?? { loadedCount: 0, unreadCount: 0, attachmentMailCount: 0 })
const profileGroups = computed(() => profile.value?.groups ?? [])
const groupPeople = computed(() => profile.value?.sameGroupPeople ?? [])

function contactTitle(contact: { displayName: string; smtpAddress: string }) {
  return contact.displayName || contact.smtpAddress || '(unknown)'
}

function storeLabel(store: OutlookStoreDto) {
  return store.displayName || store.rootFolderPath || store.storeId || '(store)'
}

function isSelfContact(contact: AddressBookContactDto) {
  const selfEmail = profile.value?.smtpAddress?.trim().toLowerCase() || ''
  return contact.isLikelySelf || (!!selfEmail && contact.smtpAddress.trim().toLowerCase() === selfEmail)
}

async function loadProfile(force = false) {
  if (loadingProfile.value || (!force && profile.value)) return
  if (!props.dashboard.outlookFirstLoadCompleted.value) return
  loadingProfile.value = true
  profileMessage.value = '正在載入個人資訊...'
  try {
    profile.value = await outlookApi.getProfile()
    profileMessage.value = profile.value.message || '個人資訊已載入。'
  } catch (error) {
    profile.value = null
    profileMessage.value = error instanceof Error ? error.message : '個人資訊載入失敗。'
  } finally {
    loadingProfile.value = false
  }
}

onMounted(() => {
  void loadProfile()
})

watch(
  () => props.dashboard.outlookFirstLoadCompleted.value,
  (completed) => {
    if (completed) void loadProfile()
  },
  { immediate: true },
)
</script>

<template>
  <main class="profile-layout">
    <section class="panel profile-panel">
      <div class="panel-header">
        <div class="panel-title">
          <el-icon><UserFilled /></el-icon>
          <span>個人資訊</span>
          <el-tag effect="plain">{{ profile ? '已載入' : '待載入' }}</el-tag>
        </div>
        <el-button :icon="Refresh" :loading="loadingProfile" :disabled="loadingProfile || !dashboard.outlookFirstLoadCompleted.value" @click="loadProfile(true)">
          重新整理
        </el-button>
      </div>

      <div class="profile-status" :class="{ loading: loadingProfile }">
        <strong>{{ loadingProfile ? '個人資訊載入中...' : '個人資訊狀態' }}</strong>
        <span>{{ profileMessage }}</span>
      </div>

      <div class="profile-grid">
        <article class="profile-section primary">
          <span class="profile-label">郵箱名稱</span>
          <strong>{{ profile?.mailboxName || '尚未辨識' }}</strong>
          <small>{{ profile?.smtpAddress || '尚未取得 SMTP address' }}</small>
        </article>

        <article class="profile-section">
          <span class="profile-label">郵件概況</span>
          <div class="profile-metrics">
            <span>{{ mailStats.loadedCount }} 封郵件</span>
            <span>{{ mailStats.unreadCount }} 封未讀</span>
            <span>{{ mailStats.attachmentMailCount }} 封含附件</span>
          </div>
        </article>

        <article class="profile-section">
          <span class="profile-label">OST</span>
          <strong>{{ profile?.ostStores.length ?? 0 }}</strong>
          <small>{{ profile?.ostStores.map(storeLabel).join(', ') || '沒有偵測到 OST store' }}</small>
        </article>

        <article class="profile-section">
          <span class="profile-label">PST</span>
          <strong>{{ profile?.pstStores.length ?? 0 }}</strong>
          <small>{{ profile?.pstStores.map(storeLabel).join(', ') || '沒有偵測到 PST archive' }}</small>
        </article>
      </div>

      <div class="profile-columns member-focused">
        <section class="profile-section profile-groups-card">
          <span class="profile-label">所在 group</span>
          <div class="profile-scroll-list">
            <button
              v-for="group in profileGroups"
              :key="group.smtpAddress || group.id || group.displayName"
              class="profile-row group-row"
              type="button"
            >
              <strong>{{ contactTitle(group) }}</strong>
              <small>{{ group.smtpAddress || '-' }}</small>
              <span>{{ group.memberCount || group.memberSmtpAddresses.length }} 位成員</span>
            </button>
            <span v-if="!profileGroups.length" class="profile-empty">尚未找到所在 group。</span>
          </div>
        </section>

        <section class="profile-section profile-members-card">
          <span class="profile-label">同 group 人員</span>
          <div class="profile-scroll-list">
            <button
              v-for="person in groupPeople"
              :key="person.smtpAddress || person.id || person.displayName"
              class="profile-row member-row"
              :class="{ self: isSelfContact(person) }"
              type="button"
            >
              <span>
                <strong>{{ contactTitle(person) }}</strong>
                <el-tag v-if="isSelfContact(person)" size="small" type="warning" effect="plain">使用者</el-tag>
              </span>
              <small>{{ person.smtpAddress || '-' }}</small>
              <small>{{ person.jobTitle || person.department || 'group member' }}</small>
            </button>
            <span v-if="!groupPeople.length" class="profile-empty">尚未找到同 group 人員。</span>
          </div>
        </section>

        <section class="profile-section">
          <span class="profile-label">資料檔</span>
          <div class="profile-scroll-list">
            <button v-for="store in profile?.stores ?? []" :key="store.storeId" class="profile-row" type="button">
              <strong>{{ storeLabel(store) }}</strong>
              <small>{{ store.storeKind || 'store' }} · {{ store.storeFilePath || store.rootFolderPath || '-' }}</small>
            </button>
          </div>
        </section>
      </div>
    </section>
  </main>
</template>
