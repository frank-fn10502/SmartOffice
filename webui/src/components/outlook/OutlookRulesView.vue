<script setup lang="ts">
import { Delete, Edit, Plus, Refresh, Select } from '@element-plus/icons-vue'
import type { OutlookDashboardState } from '../../composables/useOutlookDashboard'
import type { OutlookRuleDto } from '../../models/outlook'

const props = defineProps<{
  dashboard: OutlookDashboardState
}>()

const {
  categories,
  deleteRule,
  editRule,
  folderOptions,
  loadingRules,
  operationLoading,
  outlookBusy,
  requestRules,
  resetRuleDraft,
  ruleDraft,
  ruleDraftIsEditing,
  rules,
  saveRule,
  toggleRuleEnabled,
} = props.dashboard

function ruleTone(rule: OutlookRuleDto) {
  if (!rule.enabled) return 'info'
  if (!rule.canModifyDefinition) return 'warning'
  return 'success'
}

function ruleState(rule: OutlookRuleDto) {
  if (!rule.enabled) return '停用'
  if (!rule.canModifyDefinition) return '部分可改'
  return '啟用'
}
</script>

<template>
  <main class="rules-layout">
    <section class="rules-list-panel">
      <div class="panel-header">
        <div>
          <h2>Outlook Rules</h2>
          <p>{{ rules.length }} rules</p>
        </div>
        <div class="panel-actions">
          <el-button :icon="Plus" :disabled="outlookBusy" @click="resetRuleDraft()">新增</el-button>
          <el-button :icon="Refresh" :loading="loadingRules" :disabled="outlookBusy && !loadingRules" @click="requestRules()">
            同步
          </el-button>
        </div>
      </div>

      <el-table
        v-loading="loadingRules"
        :data="rules"
        height="calc(100vh - 208px)"
        empty-text="尚未同步 Outlook rules"
        class="rules-table"
      >
        <el-table-column prop="executionOrder" label="#" width="64" />
        <el-table-column label="Rule" min-width="220">
          <template #default="{ row }">
            <div class="rule-name-cell">
              <strong>{{ row.name }}</strong>
              <el-tag :type="ruleTone(row)" effect="plain">{{ ruleState(row) }}</el-tag>
            </div>
            <div class="rule-meta">{{ row.ruleType }}{{ row.isLocalRule ? ' · local' : '' }}</div>
          </template>
        </el-table-column>
        <el-table-column label="Conditions" min-width="260">
          <template #default="{ row }">
            <div class="rule-chip-row">
              <el-tag v-for="condition in row.conditions" :key="condition" effect="plain">
                {{ condition }}
              </el-tag>
            </div>
          </template>
        </el-table-column>
        <el-table-column label="Actions" min-width="260">
          <template #default="{ row }">
            <div class="rule-chip-row">
              <el-tag v-for="action in row.actions" :key="action" type="success" effect="plain">
                {{ action }}
              </el-tag>
            </div>
          </template>
        </el-table-column>
        <el-table-column label="操作" width="190" fixed="right">
          <template #default="{ row, $index }">
            <div class="rule-row-actions">
              <el-switch
                :model-value="row.enabled"
                :disabled="outlookBusy"
                @change="(value: boolean | string | number) => toggleRuleEnabled(row, Boolean(value))"
              />
              <el-button :icon="Edit" circle :disabled="outlookBusy || !row.canModifyDefinition" @click="editRule($index)" />
              <el-button :icon="Delete" circle type="danger" :disabled="outlookBusy" @click="deleteRule(row)" />
            </div>
          </template>
        </el-table-column>
      </el-table>
    </section>

    <aside class="rule-editor-panel">
      <div class="panel-header compact">
        <div>
          <h2>{{ ruleDraftIsEditing ? '修改 Rule' : '新增 Rule' }}</h2>
          <p>receive rule · supported conditions/actions</p>
        </div>
      </div>

      <el-form label-position="top" class="rule-form" @submit.prevent>
        <el-form-item label="名稱">
          <el-input v-model="ruleDraft.ruleName" :disabled="outlookBusy" maxlength="128" />
        </el-form-item>

        <div class="rule-form-grid">
          <el-form-item label="類型">
            <el-segmented
              v-model="ruleDraft.ruleType"
              :options="[
                { label: '接收', value: 'receive' },
                { label: '寄出', value: 'send' },
              ]"
              :disabled="outlookBusy"
            />
          </el-form-item>
          <el-form-item label="狀態">
            <el-switch v-model="ruleDraft.enabled" :disabled="outlookBusy" active-text="啟用" inactive-text="停用" />
          </el-form-item>
        </div>

        <el-divider content-position="left">條件</el-divider>
        <el-form-item label="Subject contains">
          <el-input v-model="ruleDraft.subjectContains" :disabled="outlookBusy" placeholder="用逗號或換行分隔" />
        </el-form-item>
        <el-form-item label="Body contains">
          <el-input v-model="ruleDraft.bodyContains" :disabled="outlookBusy" placeholder="用逗號或換行分隔" />
        </el-form-item>
        <el-form-item label="Sender address contains">
          <el-input v-model="ruleDraft.senderAddressContains" :disabled="outlookBusy" placeholder="example.com" />
        </el-form-item>
        <el-form-item label="Category">
          <el-select v-model="ruleDraft.categories" multiple filterable :disabled="outlookBusy" placeholder="選擇分類">
            <el-option v-for="category in categories" :key="category.name" :label="category.name" :value="category.name" />
          </el-select>
        </el-form-item>
        <el-form-item label="Attachment">
          <el-segmented
            v-model="ruleDraft.hasAttachment"
            :options="[
              { label: '不限', value: 'any' },
              { label: '有附件', value: 'yes' },
              { label: '無附件', value: 'no' },
            ]"
            :disabled="outlookBusy"
          />
        </el-form-item>

        <el-divider content-position="left">動作</el-divider>
        <el-form-item label="Move to folder">
          <el-select v-model="ruleDraft.moveToFolderPath" filterable clearable :disabled="outlookBusy" placeholder="不移動">
            <el-option v-for="folder in folderOptions" :key="folder.folderPath" :label="folder.label" :value="folder.folderPath" />
          </el-select>
        </el-form-item>
        <el-form-item label="Assign categories">
          <el-select v-model="ruleDraft.assignCategories" multiple filterable :disabled="outlookBusy" placeholder="不設定分類">
            <el-option v-for="category in categories" :key="category.name" :label="category.name" :value="category.name" />
          </el-select>
        </el-form-item>
        <el-checkbox v-model="ruleDraft.markAsTask" :disabled="outlookBusy">Mark as task</el-checkbox>
        <el-checkbox v-model="ruleDraft.stopProcessingMoreRules" :disabled="outlookBusy">Stop processing more rules</el-checkbox>

        <div class="rule-editor-actions">
          <el-button :disabled="outlookBusy" @click="resetRuleDraft()">清除</el-button>
          <el-button type="primary" :icon="Select" :loading="operationLoading" :disabled="outlookBusy" @click="saveRule">
            儲存
          </el-button>
        </div>
      </el-form>
    </aside>
  </main>
</template>
