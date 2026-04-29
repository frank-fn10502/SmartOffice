<script setup lang="ts">
import { Monitor, Refresh, Setting } from '@element-plus/icons-vue'
import { ElAlert, ElButton, ElCard, ElConfigProvider, ElIcon, ElSpace, ElTag } from 'element-plus'
</script>

<template>
  <el-config-provider size="default">
    <main class="app-shell">
      <section class="toolbar">
        <div>
          <p class="eyebrow">SmartOffice.Hub Web UI</p>
          <h1>Vue workspace is ready</h1>
        </div>

        <el-space wrap>
          <el-tag type="success" effect="light">Vue 3</el-tag>
          <el-tag type="info" effect="light">Vite</el-tag>
          <el-tag type="warning" effect="light">Element Plus</el-tag>
        </el-space>
      </section>

      <el-card shadow="never" class="status-card">
        <template #header>
          <div class="card-title">
            <el-icon><Monitor /></el-icon>
            <span>Local build baseline</span>
          </div>
        </template>

        <el-alert
          title="這是前端 toolchain 的最小入口，尚未取代既有 dashboard。"
          type="info"
          :closable="false"
          show-icon
        />

        <div class="actions">
          <el-button type="primary" :icon="Refresh">Build Preview</el-button>
          <el-button :icon="Setting">Settings</el-button>
        </div>
      </el-card>
    </main>
  </el-config-provider>
</template>
