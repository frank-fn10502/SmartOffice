<script setup lang="ts">
import { computed, onMounted, ref, watch } from 'vue'
import { Refresh, UserFilled } from '@element-plus/icons-vue'
import { outlookApi } from '../api/outlook'
import type { OutlookDashboardState } from '../composables/useOutlookDashboard'
import type { OutlookProfileDto, OutlookProfileGroupNodeDto, OutlookStoreDto } from '../models/outlook'

const props = defineProps<{
  dashboard: OutlookDashboardState
}>()

const profile = ref<OutlookProfileDto | null>(null)
const loadingProfile = ref(false)
const profileMessage = ref('等待 Outlook 郵件載入完成。')

const mailStats = computed(() => profile.value?.mailStats ?? { loadedCount: 0, unreadCount: 0, attachmentMailCount: 0 })

type ProfileTreeNode = {
  key: string
  label: string
  email: string
  children: ProfileTreeNode[]
}

const profileGroupTreeProps = {
  children: 'children',
}

const profileGroupTreeNodes = computed(() => (profile.value?.groupTree ?? []).map(toProfileTreeNode))

function contactTitle(contact: { displayName: string; smtpAddress: string }) {
  return contact.displayName || contact.smtpAddress || '(unknown)'
}

function storeLabel(store: OutlookStoreDto) {
  return store.displayName || store.rootFolderPath || store.storeId || '(store)'
}

function toProfileTreeNode(node: OutlookProfileGroupNodeDto): ProfileTreeNode {
  const contact = node.contact
  return {
    key: contact.smtpAddress || contact.id || contact.displayName,
    label: contactTitle(contact),
    email: contact.smtpAddress || '',
    children: node.children.map(toProfileTreeNode),
  }
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
          <span class="profile-label">Mail snapshot</span>
          <div class="profile-metrics">
            <span>{{ mailStats.loadedCount }} loaded</span>
            <span>{{ mailStats.unreadCount }} unread</span>
            <span>{{ mailStats.attachmentMailCount }} attachments</span>
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

      <div class="profile-columns">
        <section class="profile-section">
          <span class="profile-label">所屬 group</span>
          <div class="profile-scroll-list">
            <el-tree
              v-if="profileGroupTreeNodes.length"
              class="profile-group-tree"
              :data="profileGroupTreeNodes"
              node-key="key"
              default-expand-all
              :props="profileGroupTreeProps"
            >
              <template #default="{ data }">
                <span class="profile-tree-row">
                  <strong>{{ data.label }}</strong>
                  <small>{{ data.email || '-' }}</small>
                </span>
              </template>
            </el-tree>
            <span v-if="!profileGroupTreeNodes.length" class="profile-empty">尚未從 Hub 個人資訊取得 group 階層。</span>
          </div>
        </section>

        <section class="profile-section">
          <span class="profile-label">同 group 人員</span>
          <div class="profile-scroll-list">
            <button v-for="person in profile?.sameGroupPeople ?? []" :key="person.smtpAddress || person.id" class="profile-row" type="button">
              <strong>{{ contactTitle(person) }}</strong>
              <small>{{ person.smtpAddress || '-' }}</small>
            </button>
            <span v-if="!profile?.sameGroupPeople.length" class="profile-empty">尚未從 Hub 個人資訊取得同 group 人員。</span>
          </div>
        </section>

        <section class="profile-section">
          <span class="profile-label">Stores</span>
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
