<script setup lang="ts">
import { computed, ref } from 'vue'
import { Connection, CopyDocument, FolderOpened, Operation, Refresh, Setting, Warning } from '@element-plus/icons-vue'
import { ElMessage } from 'element-plus'
import type { OutlookDashboardState } from '../features/outlook/composables/useOutlookDashboard'
import { formatTime } from '../features/outlook/utils/formatters'

const props = defineProps<{
  dashboard: OutlookDashboardState
}>()

const {
  addinLogs,
  addinStatus,
  attachmentExportRootDraft,
  attachmentExportSettings,
  loadingSignalRPing,
  refreshAdminData,
  requestSignalRPing,
  resetAttachmentExportRoot,
  saveAttachmentExportSettings,
  savingAttachmentExportSettings,
} = props.dashboard

const adminSection = ref<'runtime' | 'settings'>('runtime')
const healthTone = computed(() => addinStatus.value.connected ? 'online' : 'offline')
const healthLabel = computed(() => addinStatus.value.connected ? 'Online' : 'Offline')
const lastCommandLabel = computed(() => addinStatus.value.lastCommand || '-')
const logCounts = computed(() => {
  return addinLogs.value.reduce(
    (counts, log) => {
      const level = log.level.toLowerCase()
      if (level === 'error') counts.error += 1
      else if (level === 'warn' || level === 'warning') counts.warn += 1
      else counts.info += 1
      return counts
    },
    { error: 0, warn: 0, info: 0 },
  )
})

async function copyLogs() {
  if (addinLogs.value.length === 0) {
    ElMessage.warning('目前沒有 log 可複製')
    return
  }
  const text = addinLogs.value
    .map((log) => `[${formatTime(log.timestamp)}] ${log.level.toUpperCase()} ${log.message}`)
    .join('\n')
  try {
    await navigator.clipboard.writeText(text)
    ElMessage.success(`已複製 ${addinLogs.value.length} 筆 log`)
  } catch {
    ElMessage.error('無法複製 log，瀏覽器未開放 clipboard 權限')
  }
}
</script>

<template>
  <main class="admin-layout">
    <section class="admin-hero panel">
      <div class="admin-hero-main">
        <div class="admin-kicker">Outlook Runtime</div>
        <h1>Admin</h1>
        <div class="admin-health-line">
          <span class="admin-health-dot" :class="healthTone" />
          <span>Outlook connection {{ healthLabel }}</span>
          <span>Last command: {{ lastCommandLabel }}</span>
        </div>
      </div>
      <div class="admin-hero-actions">
        <el-button :icon="Refresh" @click="refreshAdminData">Refresh</el-button>
        <el-button type="primary" :loading="loadingSignalRPing" :disabled="!addinStatus.connected" @click="requestSignalRPing">
          Connection Ping
        </el-button>
      </div>
    </section>

    <section class="admin-workbench">
      <aside class="panel admin-nav">
        <button
          class="admin-nav-item"
          :class="{ active: adminSection === 'runtime' }"
          type="button"
          @click="adminSection = 'runtime'"
        >
          <el-icon><Operation /></el-icon>
          <span>
            <strong>Runtime</strong>
            <small>Logs and connection</small>
          </span>
        </button>
        <button
          class="admin-nav-item"
          :class="{ active: adminSection === 'settings' }"
          type="button"
          @click="adminSection = 'settings'"
        >
          <el-icon><Setting /></el-icon>
          <span>
            <strong>Settings</strong>
            <small>Configuration</small>
          </span>
        </button>
      </aside>

      <section v-if="adminSection === 'runtime'" class="admin-runtime-view">
        <section class="panel admin-connection-panel">
          <div class="admin-card-header">
            <div class="panel-title">
              <el-icon><Connection /></el-icon>
              <span>Connection</span>
            </div>
            <el-tag :type="addinStatus.connected ? 'success' : 'danger'" effect="plain">{{ healthLabel }}</el-tag>
          </div>

          <div class="admin-metric-stack">
            <div class="admin-metric primary">
              <span>Outlook connection</span>
              <strong :class="healthTone">{{ healthLabel }}</strong>
            </div>
            <div class="admin-metric">
              <span>Last Connect</span>
              <strong>{{ formatTime(addinStatus.lastPollTime) }}</strong>
            </div>
            <div class="admin-metric">
              <span>Last Push</span>
              <strong>{{ formatTime(addinStatus.lastPushTime) }}</strong>
            </div>
            <div class="admin-metric">
              <span>Last Command</span>
              <strong>{{ lastCommandLabel }}</strong>
            </div>
          </div>
        </section>

        <section class="panel admin-logs-panel">
          <div class="admin-card-header">
            <div class="panel-title">
              <el-icon><Operation /></el-icon>
              <span>Runtime Logs</span>
            </div>
            <div class="admin-log-stats">
              <el-tag effect="plain">{{ logCounts.info }} info</el-tag>
              <el-tag type="warning" effect="plain">{{ logCounts.warn }} warn</el-tag>
              <el-tag type="danger" effect="plain">{{ logCounts.error }} error</el-tag>
              <el-button :icon="CopyDocument" size="small" :disabled="addinLogs.length === 0" @click="copyLogs">
                Copy logs
              </el-button>
            </div>
          </div>

          <div class="logs">
            <div v-if="addinLogs.length === 0" class="admin-empty-log">
              <el-icon><FolderOpened /></el-icon>
              <span>No logs yet.</span>
            </div>
            <div v-for="(log, index) in addinLogs" :key="`${log.timestamp}-${index}`" class="log-entry" :class="log.level.toLowerCase()">
              <span class="log-time">{{ formatTime(log.timestamp) }}</span>
              <span class="log-level">
                <el-icon v-if="log.level.toLowerCase() === 'error'"><Warning /></el-icon>
                {{ log.level.toUpperCase() }}
              </span>
              <span class="log-message">{{ log.message }}</span>
            </div>
          </div>
        </section>
      </section>

      <section v-else class="admin-settings-layout">
        <article class="panel admin-settings-panel">
          <div class="admin-card-header">
            <div class="panel-title">
              <el-icon><Setting /></el-icon>
              <span>Settings</span>
            </div>
          </div>

          <div class="admin-setting-section">
            <strong>Attachment Export</strong>
            <div class="admin-path-summary">
              <div>
                <span>Current Root</span>
                <strong>{{ attachmentExportSettings.rootPath || '載入中...' }}</strong>
              </div>
              <div>
                <span>Default Root</span>
                <strong>{{ attachmentExportSettings.defaultRootPath || '載入中...' }}</strong>
              </div>
            </div>

            <div class="admin-setting-form">
              <div class="inspector-field">
                <span>Export root</span>
                <el-input v-model="attachmentExportRootDraft" :placeholder="attachmentExportSettings.defaultRootPath || '$HOME/SmartOffice/Attachments'" />
              </div>
              <div class="field-hint">
                macOS / Linux 預設會放在使用者 home 底下的 SmartOffice/Attachments；Windows 會依序使用 E:\、D:\、C:\ 底下的 SmartOffice\Attachments。
              </div>
              <div class="admin-actions">
                <el-button type="primary" :loading="savingAttachmentExportSettings" @click="saveAttachmentExportSettings">
                  儲存
                </el-button>
                <el-button :disabled="savingAttachmentExportSettings" @click="resetAttachmentExportRoot">
                  使用預設
                </el-button>
              </div>
            </div>
          </div>
        </article>
      </section>
    </section>
  </main>
</template>
