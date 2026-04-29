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

interface FolderDto {
  name: string
  folderPath: string
  itemCount: number
  subFolders: FolderDto[]
}

const props = defineProps<{
  folder: FolderDto
  level: number
  expandedFolders: Set<string>
  selectedFolderPath: string
}>()

const emit = defineEmits<{
  toggle: [path: string]
  select: [path: string]
}>()

const hiddenFolderNames = [
  'common views',
  'finder',
  'reminders',
  'quick step',
  'conversation history',
  'conversation action',
  'server failures',
  'local failures',
  'conflicts',
  'sync issues',
  'rss',
  'social network',
  'people',
  'externalcontacts',
  'yammer',
]

function isHiddenFolder(name: string) {
  const lowerName = name.toLowerCase()
  return hiddenFolderNames.some((hidden) => lowerName.includes(hidden))
}

function visibleChildren(folder: FolderDto) {
  return (folder.subFolders ?? []).filter((child) => !isHiddenFolder(child.name))
}

function folderType(name: string) {
  const lowerName = name.toLowerCase()
  if (lowerName === 'inbox') return 'inbox'
  if (lowerName === 'sent items' || lowerName.includes('sent')) return 'sent'
  if (lowerName === 'drafts') return 'drafts'
  if (lowerName === 'deleted items' || lowerName.includes('deleted')) return 'deleted'
  if (lowerName === 'junk email' || lowerName === 'junk e-mail') return 'junk'
  if (lowerName === 'archive') return 'archive'
  if (lowerName === 'outbox') return 'outbox'
  return 'normal'
}

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
      :class="[folderType(folder.name), { selected: selectedFolderPath === folder.folderPath }]"
      :style="{ paddingLeft: `${level * 16 + 6}px` }"
      @click="emit('select', folder.folderPath)"
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

    <div v-if="expandedFolders.has(folder.folderPath)" class="folder-children">
      <FolderNode
        v-for="child in visibleChildren(folder)"
        :key="child.folderPath"
        :folder="child"
        :level="level + 1"
        :expanded-folders="expandedFolders"
        :selected-folder-path="selectedFolderPath"
        @toggle="emit('toggle', $event)"
        @select="emit('select', $event)"
      />
    </div>
  </div>
</template>
