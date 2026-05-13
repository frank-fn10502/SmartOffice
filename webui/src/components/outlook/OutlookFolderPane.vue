<script setup lang="ts">
import { Folder, Refresh } from '@element-plus/icons-vue'
import FolderNode from '../FolderNode.vue'
import type { OutlookDashboardState } from '../../composables/useOutlookDashboard'

const props = defineProps<{
  dashboard: OutlookDashboardState
}>()

const {
  cancelCreateFolder,
  createFolder,
  creatingFolderName,
  creatingFolderParentPath,
  dragOverFolderPath,
  draggedMailId,
  expandedFolders,
  folderStores,
  loadingFolders,
  moveDraggedMail,
  openFolderContextMenu,
  outlookBusy,
  requestFolders,
  selectFolder,
  selectedFolderPath,
  setDragOverFolder,
  toggleFolder,
  visibleFolders,
} = props.dashboard
</script>

<template>
  <section class="panel outlook-folder-pane">
    <div class="panel-header">
      <div class="panel-title">
        <el-icon><Folder /></el-icon>
        <span>Folders</span>
      </div>
      <el-button
        :icon="Refresh"
        circle
        :loading="loadingFolders"
        :disabled="outlookBusy && !loadingFolders"
        @click="requestFolders"
      />
    </div>

    <div class="folder-list outlook-folder-list">
      <p v-if="visibleFolders.length === 0 && !loadingFolders" class="hint">Waiting for folders...</p>
      <FolderNode
        v-for="folder in visibleFolders"
        :key="folder.folderPath"
        :folder="folder"
        :store="folderStores.find((store) => store.storeId === folder.storeId)"
        :level="0"
        :expanded-folders="expandedFolders"
        :selected-folder-path="selectedFolderPath"
        :creating-folder-parent-path="creatingFolderParentPath"
        :creating-folder-name="creatingFolderName"
        :folder-busy="outlookBusy"
        :can-drop-mail="Boolean(draggedMailId) && !outlookBusy"
        :active-drop-folder-path="dragOverFolderPath"
        @toggle="toggleFolder"
        @select="selectFolder"
        @context="openFolderContextMenu"
        @update:creating-folder-name="creatingFolderName = $event"
        @create="createFolder($event.parentPath, $event.name)"
        @cancel-create="cancelCreateFolder"
        @drag-mail-over="setDragOverFolder"
        @drop-mail="moveDraggedMail"
      />
      <div v-if="loadingFolders" class="pane-loading">
        <span>Outlook folder 同步中...</span>
      </div>
    </div>
  </section>
</template>
