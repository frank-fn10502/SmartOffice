<script setup lang="ts">
import {
  ArrowRight,
  Box,
  Delete,
  Folder as FolderIcon,
  Message,
  Promotion,
  Tickets,
  Warning,
} from '@element-plus/icons-vue'
import type { FolderDto } from '../models/outlook'
import { folderType, visibleChildren } from '../utils/folders'

const props = defineProps<{
  folder: FolderDto
  level: number
  expandedFolders: Set<string>
  selectedFolderPath: string
  creatingFolderParentPath: string
  creatingFolderName: string
  folderBusy: boolean
  canDropMail: boolean
  activeDropFolderPath: string
}>()

const emit = defineEmits<{
  toggle: [path: string]
  select: [path: string]
  context: [payload: { path: string; x: number; y: number }]
  'update:creatingFolderName': [name: string]
  create: [payload: { parentPath: string; name: string }]
  cancelCreate: []
  dragMailOver: [folderPath: string]
  dropMail: [folderPath: string]
}>()

function folderIcon(name: string) {
  const icons = {
    inbox: Message,
    sent: Promotion,
    drafts: Tickets,
    deleted: Delete,
    junk: Warning,
    archive: Box,
    outbox: Promotion,
    normal: FolderIcon,
  }
  return icons[folderType(name)]
}
</script>

<template>
  <div class="folder-node">
    <div
      class="folder-row"
      :class="[
        folderType(folder.name),
        {
          selected: selectedFolderPath === folder.folderPath,
          'drop-target': canDropMail && !folderBusy,
          'drop-active': activeDropFolderPath === folder.folderPath,
        },
      ]"
      :style="{ paddingLeft: `${level * 16 + 6}px` }"
      @click="emit('select', folder.folderPath)"
      @contextmenu.prevent="emit('context', { path: folder.folderPath, x: $event.clientX, y: $event.clientY })"
      @dragenter.prevent.stop="!folderBusy && emit('dragMailOver', folder.folderPath)"
      @dragover.prevent.stop="!folderBusy && emit('dragMailOver', folder.folderPath)"
      @dragleave.stop="emit('dragMailOver', '')"
      @drop.prevent.stop="!folderBusy && emit('dropMail', folder.folderPath)"
    >
      <button
        class="folder-toggle"
        :class="{ expanded: expandedFolders.has(folder.folderPath), empty: visibleChildren(folder).length === 0 }"
        type="button"
        @click.stop="emit('toggle', folder.folderPath)"
      >
        <el-icon v-if="visibleChildren(folder).length > 0"><ArrowRight /></el-icon>
      </button>

      <el-icon class="folder-kind">
        <component :is="folderIcon(folder.name)" />
      </el-icon>
      <span class="folder-name">{{ folder.name }}</span>
      <span class="folder-count">{{ folder.itemCount }}</span>
    </div>

    <div
      v-if="creatingFolderParentPath === folder.folderPath"
      class="folder-inline-create"
      :style="{ paddingLeft: `${(level + 1) * 16 + 24}px` }"
      @click.stop
    >
      <el-input
        :model-value="creatingFolderName"
        size="small"
        autofocus
        placeholder="New folder"
        :disabled="folderBusy"
        @update:model-value="emit('update:creatingFolderName', $event)"
        @keydown.enter.prevent="emit('create', { parentPath: folder.folderPath, name: creatingFolderName })"
        @keydown.esc.prevent="emit('cancelCreate')"
        @blur="emit('cancelCreate')"
      />
    </div>

    <div v-if="expandedFolders.has(folder.folderPath)" class="folder-children">
      <FolderNode
        v-for="child in visibleChildren(folder)"
        :key="child.folderPath"
        :folder="child"
        :level="level + 1"
        :expanded-folders="expandedFolders"
        :selected-folder-path="selectedFolderPath"
        :creating-folder-parent-path="creatingFolderParentPath"
        :creating-folder-name="creatingFolderName"
        :folder-busy="folderBusy"
        :can-drop-mail="canDropMail"
        :active-drop-folder-path="activeDropFolderPath"
        @toggle="emit('toggle', $event)"
        @select="emit('select', $event)"
        @context="emit('context', $event)"
        @update:creating-folder-name="emit('update:creatingFolderName', $event)"
        @create="emit('create', $event)"
        @cancel-create="emit('cancelCreate')"
        @drag-mail-over="emit('dragMailOver', $event)"
        @drop-mail="emit('dropMail', $event)"
      />
    </div>
  </div>
</template>
