<script setup lang="ts">
import { computed } from 'vue'
import { Connection, FolderOpened, Operation, Refresh, Setting, Warning } from '@element-plus/icons-vue'
import type { OutlookDashboardState } from '../composables/useOutlookDashboard'
import { formatTime } from '../utils/formatters'

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
</script>

<template>
  <main class="admin-layout">
    <section class="admin-hero panel">
      <div class="admin-hero-main">
        <div class="admin-kicker">Runtime Console</div>
        <h1>Hub Admin</h1>
        <div class="admin-health-line">
          <span class="admin-health-dot" :class="healthTone" />
          <span>Outlook Add-in {{ healthLabel }}</span>
          <span>Last command: {{ lastCommandLabel }}</span>
        </div>
      </div>
      <div class="admin-hero-actions">
        <el-button :icon="Refresh" @click="refreshAdminData">Refresh</el-button>
        <el-button type="primary" :loading="loadingSignalRPing" :disabled="!addinStatus.connected" @click="requestSignalRPing">
          SignalR Ping
        </el-button>
      </div>
    </section>

    <section class="admin-grid">
      <article class="panel admin-card">
        <div class="admin-card-header">
          <div class="panel-title">
            <el-icon><Connection /></el-icon>
            <span>Connection</span>
          </div>
          <el-tag :type="addinStatus.connected ? 'success' : 'danger'" effect="plain">{{ healthLabel }}</el-tag>
        </div>

        <div class="admin-metric-stack">
          <div class="admin-metric primary">
            <span>Outlook Add-in</span>
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
      </article>

      <article class="panel admin-card">
        <div class="admin-card-header">
          <div class="panel-title">
            <el-icon><Setting /></el-icon>
            <span>Attachment Export</span>
          </div>
        </div>

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
      </article>
    </section>

    <section class="panel admin-logs-panel">
      <div class="admin-card-header">
        <div class="panel-title">
          <el-icon><Operation /></el-icon>
          <span>Add-in Logs</span>
        </div>
        <div class="admin-log-stats">
          <el-tag effect="plain">{{ logCounts.info }} info</el-tag>
          <el-tag type="warning" effect="plain">{{ logCounts.warn }} warn</el-tag>
          <el-tag type="danger" effect="plain">{{ logCounts.error }} error</el-tag>
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
  </main>
</template>
