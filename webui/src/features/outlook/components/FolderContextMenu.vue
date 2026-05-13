<script setup lang="ts">
import type { OutlookDashboardState } from '../composables/useOutlookDashboard'

const props = defineProps<{
  dashboard: OutlookDashboardState
}>()

const {
  contextFolderName,
  createFolderFromContext,
  deleteFolderFromContext,
  fetchMailsFromContext,
  folderContextMenu,
  outlookBusy,
} = props.dashboard
</script>

<template>
  <div
    v-if="folderContextMenu.visible"
    class="folder-context-menu"
    :style="{ left: `${folderContextMenu.x}px`, top: `${folderContextMenu.y}px` }"
    @click.stop
  >
    <div class="context-menu-title">{{ contextFolderName }}</div>
    <button type="button" :disabled="outlookBusy" @click="fetchMailsFromContext">抓取郵件</button>
    <button type="button" :disabled="outlookBusy" @click="createFolderFromContext">新增子資料夾</button>
    <button class="danger" type="button" :disabled="outlookBusy" @click="deleteFolderFromContext">刪除此資料夾</button>
  </div>
</template>
