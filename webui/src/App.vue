<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { Monitor, Moon, Platform, Sunny } from '@element-plus/icons-vue'
import { useTheme } from './composables/useTheme'
import { useOutlookDashboard } from './features/outlook/composables/useOutlookDashboard'
import type { HubPage } from './features/outlook/models/outlook'
import AdminPage from './views/AdminPage.vue'
import OutlookPage from './views/OutlookPage.vue'

const dashboard = useOutlookDashboard()
const { isDarkTheme, themeLabel, toggleTheme } = useTheme()
const activePage = ref<HubPage>('outlook')
const signalRState = computed(() => dashboard.signalRState.value)

const pageOptions = computed(() => [
  { label: 'Outlook', value: 'outlook' },
  { label: 'Admin', value: 'admin' },
  { label: 'Swagger', value: 'swagger' },
])

watch(activePage, async (page) => {
  if (page !== 'admin') return
  await Promise.allSettled([
    dashboard.refreshAdminData(),
    dashboard.loadAttachmentExportSettings(),
  ])
})
</script>

<template>
  <el-config-provider size="default">
    <div class="app-shell">
      <header class="topbar">
        <div class="brand">
          <el-icon><Monitor /></el-icon>
          <span>SmartOffice Hub</span>
          <el-tag :type="signalRState === 'connected' ? 'success' : 'danger'" effect="dark">
            {{ signalRState }}
          </el-tag>
        </div>

        <nav class="nav-actions">
          <el-button
            class="theme-toggle"
            :icon="isDarkTheme ? Sunny : Moon"
            @click="toggleTheme"
          >
            {{ themeLabel }}
          </el-button>
          <el-segmented
            :model-value="activePage"
            :options="pageOptions"
            @update:model-value="(value: string | number | boolean) => activePage = value as HubPage"
          />
        </nav>
      </header>

      <OutlookPage v-if="activePage === 'outlook'" :dashboard="dashboard" />
      <AdminPage v-else-if="activePage === 'admin'" :dashboard="dashboard" />
      <main v-else class="swagger-layout">
        <section class="panel swagger-panel">
          <div class="panel-header">
            <div class="panel-title">
              <el-icon><Platform /></el-icon>
              <span>Swagger</span>
            </div>
            <el-link href="/swagger/index.html" target="_blank" type="primary">Open</el-link>
          </div>
          <iframe class="swagger-frame" src="/swagger/index.html" title="Swagger" />
        </section>
      </main>
    </div>
  </el-config-provider>
</template>
