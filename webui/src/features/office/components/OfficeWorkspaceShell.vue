<script setup lang="ts">
import type { Component } from 'vue'
import type { WorkspaceNavOption, WorkspaceNavValue } from '../models/workspace'

defineProps<{
  title: string
  icon: Component
  statusLabel?: string
  statusType?: 'success' | 'warning' | 'info' | 'primary' | 'danger'
  activeView: WorkspaceNavValue
  navOptions: WorkspaceNavOption[]
}>()

defineEmits<{
  updateActiveView: [value: WorkspaceNavValue]
}>()
</script>

<template>
  <div class="office-workspace">
    <div class="feature-toolbar">
      <div class="feature-title">
        <el-icon><component :is="icon" /></el-icon>
        <span>{{ title }}</span>
        <el-tag v-if="statusLabel" :type="statusType" effect="plain">
          {{ statusLabel }}
        </el-tag>
      </div>

      <el-segmented
        :model-value="activeView"
        :options="navOptions"
        @update:model-value="$emit('updateActiveView', $event)"
      />
    </div>

    <slot />
  </div>
</template>
