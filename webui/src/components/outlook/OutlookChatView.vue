<script setup lang="ts">
import { ChatDotRound } from '@element-plus/icons-vue'
import type { OutlookDashboardState } from '../../composables/useOutlookDashboard'
import { formatTime } from '../../utils/formatters'

const props = defineProps<{
  dashboard: OutlookDashboardState
}>()

const {
  chatMessages,
  chatPanelRef,
  chatText,
  sendChat,
} = props.dashboard
</script>

<template>
  <main class="chat-layout">
    <section class="panel chat-page-panel">
      <div class="panel-header">
        <div class="panel-title">
          <el-icon><ChatDotRound /></el-icon>
          <span>Chat</span>
        </div>
      </div>

      <div ref="chatPanelRef" class="chat-messages">
        <div
          v-for="(message, index) in chatMessages"
          :key="message.id ?? `${message.timestamp}-${index}`"
          class="chat-message"
          :class="{ web: message.source === 'web' }"
        >
          <span class="chat-meta">[{{ message.source }}] {{ formatTime(message.timestamp) }}</span>
          <span class="chat-bubble">{{ message.text }}</span>
        </div>
      </div>

      <div class="chat-input">
        <el-input v-model="chatText" placeholder="Send message..." @keydown.enter="sendChat" />
        <el-button type="primary" @click="sendChat">Send</el-button>
      </div>
    </section>
  </main>
</template>
