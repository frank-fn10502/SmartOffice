<script setup lang="ts">
import { Delete, Edit, Plus, Refresh, Select } from '@element-plus/icons-vue'
import { computed, ref } from 'vue'
import type { OutlookDashboardState } from '../composables/useOutlookDashboard'
import type { OutlookRuleDto } from '../models/outlook'

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
  selectedRuleIndex,
  toggleRuleEnabled,
} = props.dashboard

const ruleEditorVisible = ref(false)
const ruleNameMissing = computed(() => !ruleDraft.value.ruleName.trim())
const ruleHasCondition = computed(() => Boolean(
  ruleDraft.value.subjectContains.trim()
    || ruleDraft.value.bodyContains.trim()
    || ruleDraft.value.bodyOrSubjectContains.trim()
    || ruleDraft.value.messageHeaderContains.trim()
    || ruleDraft.value.senderAddressContains.trim()
    || ruleDraft.value.recipientAddressContains.trim()
    || ruleDraft.value.categories.length
    || ruleDraft.value.hasAttachment !== 'any'
    || ruleDraft.value.importance !== 'any'
    || ruleDraft.value.toMe
    || ruleDraft.value.toOrCcMe
    || ruleDraft.value.onlyToMe
    || ruleDraft.value.meetingInviteOrUpdate,
))
const ruleHasAction = computed(() => Boolean(
  ruleDraft.value.moveToFolderPath
    || ruleDraft.value.copyToFolderPath
    || ruleDraft.value.assignCategories.length
    || ruleDraft.value.clearCategories
    || ruleDraft.value.markAsTask
    || ruleDraft.value.delete
    || ruleDraft.value.desktopAlert
    || ruleDraft.value.stopProcessingMoreRules,
))
const ruleCanSave = computed(() => !outlookBusy.value && !ruleNameMissing.value && ruleHasCondition.value && ruleHasAction.value)

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

function ruleRowClassName({ rowIndex }: { rowIndex: number }) {
  return selectedRuleIndex.value === rowIndex ? 'selected-rule-row' : ''
}

function openNewRuleDialog() {
  resetRuleDraft()
  ruleEditorVisible.value = true
}

function openEditRuleDialog(index: number) {
  editRule(index)
  ruleEditorVisible.value = true
}

async function submitRule() {
  const saved = await saveRule()
  if (saved) {
    ruleEditorVisible.value = false
  }
}
</script>

<template>
  <main class="rules-layout">
    <section class="rules-list-panel">
      <div class="panel-header">
        <div>
          <h2>Outlook Rules</h2>
          <p>{{ rules.length }} rules · {{ rules.filter((rule) => rule.enabled).length }} enabled</p>
        </div>
        <div class="panel-actions">
          <el-button :icon="Plus" :disabled="outlookBusy" @click="openNewRuleDialog()">新增</el-button>
          <el-button :icon="Refresh" :loading="loadingRules" :disabled="outlookBusy" @click="requestRules()">
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
        :row-class-name="ruleRowClassName"
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
            <div class="rule-row-actions" @click.stop>
              <el-switch
                :model-value="row.enabled"
                :disabled="outlookBusy"
                @change="(value: boolean | string | number) => toggleRuleEnabled(row, Boolean(value))"
              />
              <el-button :icon="Edit" circle :disabled="outlookBusy || !row.canModifyDefinition" @click="openEditRuleDialog($index)" />
              <el-button :icon="Delete" circle type="danger" :disabled="outlookBusy" @click="deleteRule(row)" />
            </div>
          </template>
        </el-table-column>
      </el-table>
    </section>

    <el-dialog
      v-model="ruleEditorVisible"
      width="820px"
      top="5vh"
      class="rule-editor-dialog"
      destroy-on-close
    >
      <template #header>
        <div class="rule-dialog-head">
          <div>
            <h2>{{ ruleDraftIsEditing ? '修改 Rule' : '新增 Rule' }}</h2>
            <span>{{ ruleDraft.ruleType === 'send' ? 'Outgoing' : 'Incoming' }} · {{ ruleDraft.enabled ? 'Enabled' : 'Disabled' }}</span>
          </div>
          <el-tag :type="ruleDraft.enabled ? 'success' : 'info'" effect="plain">
            {{ ruleDraft.enabled ? '啟用' : '停用' }}
          </el-tag>
        </div>
      </template>

      <el-form label-position="top" class="rule-form" @submit.prevent>
        <section class="rule-editor-section rule-editor-section-primary">
          <div class="rule-section-title">
            <strong>基本設定</strong>
            <span>Rule identity</span>
          </div>
          <div class="rule-form-grid">
            <el-form-item label="名稱" class="rule-field-wide" :error="ruleNameMissing ? '請輸入 rule 名稱' : ''">
              <el-input v-model="ruleDraft.ruleName" :disabled="outlookBusy" maxlength="128" />
            </el-form-item>
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
        </section>

        <section class="rule-editor-section" :class="{ 'rule-editor-section-invalid': !ruleHasCondition }">
          <div class="rule-section-title">
            <strong>條件</strong>
            <span>Conditions</span>
          </div>
          <div class="rule-form-grid">
            <el-form-item label="Subject contains">
              <el-input v-model="ruleDraft.subjectContains" :disabled="outlookBusy" placeholder="keyword" />
            </el-form-item>
            <el-form-item label="Body contains">
              <el-input v-model="ruleDraft.bodyContains" :disabled="outlookBusy" placeholder="keyword" />
            </el-form-item>
            <el-form-item label="Subject or body contains">
              <el-input v-model="ruleDraft.bodyOrSubjectContains" :disabled="outlookBusy" placeholder="keyword" />
            </el-form-item>
            <el-form-item label="Message header contains">
              <el-input v-model="ruleDraft.messageHeaderContains" :disabled="outlookBusy" placeholder="X-header or keyword" />
            </el-form-item>
            <el-form-item label="Sender address contains">
              <el-input v-model="ruleDraft.senderAddressContains" :disabled="outlookBusy" placeholder="example.com" />
            </el-form-item>
            <el-form-item label="Recipient address contains">
              <el-input v-model="ruleDraft.recipientAddressContains" :disabled="outlookBusy" placeholder="team@example.com" />
            </el-form-item>
            <el-form-item label="Category">
              <el-select v-model="ruleDraft.categories" multiple filterable :disabled="outlookBusy" placeholder="選擇分類">
                <el-option v-for="category in categories" :key="category.name" :label="category.name" :value="category.name" />
              </el-select>
            </el-form-item>
            <el-form-item label="Importance">
              <el-select v-model="ruleDraft.importance" :disabled="outlookBusy">
                <el-option label="不限" value="any" />
                <el-option label="Low" value="low" />
                <el-option label="Normal" value="normal" />
                <el-option label="High" value="high" />
              </el-select>
            </el-form-item>
            <el-form-item label="Attachment">
              <el-segmented
                v-model="ruleDraft.hasAttachment"
                :options="[
                  { label: '不限', value: 'any' },
                  { label: '有附件', value: 'yes' },
                ]"
                :disabled="outlookBusy"
              />
            </el-form-item>
            <div class="rule-option-row rule-field-wide">
              <el-checkbox v-model="ruleDraft.toMe" :disabled="outlookBusy">寄給我</el-checkbox>
              <el-checkbox v-model="ruleDraft.toOrCcMe" :disabled="outlookBusy">To 或 CC 我</el-checkbox>
              <el-checkbox v-model="ruleDraft.onlyToMe" :disabled="outlookBusy">只寄給我</el-checkbox>
              <el-checkbox v-model="ruleDraft.meetingInviteOrUpdate" :disabled="outlookBusy">會議邀請/更新</el-checkbox>
            </div>
            <div v-if="!ruleHasCondition" class="rule-validation-hint rule-field-wide">
              請至少填寫一個條件。
            </div>
          </div>
        </section>

        <section class="rule-editor-section" :class="{ 'rule-editor-section-invalid': !ruleHasAction }">
          <div class="rule-section-title">
            <strong>動作</strong>
            <span>Actions</span>
          </div>
          <div class="rule-form-grid">
            <el-form-item label="Move to folder">
              <el-select v-model="ruleDraft.moveToFolderPath" filterable clearable :disabled="outlookBusy" placeholder="不移動">
                <el-option v-for="folder in folderOptions" :key="folder.folderPath" :label="folder.label" :value="folder.folderPath" />
              </el-select>
            </el-form-item>
            <el-form-item label="Copy to folder">
              <el-select v-model="ruleDraft.copyToFolderPath" filterable clearable :disabled="outlookBusy" placeholder="不複製">
                <el-option v-for="folder in folderOptions" :key="folder.folderPath" :label="folder.label" :value="folder.folderPath" />
              </el-select>
            </el-form-item>
            <el-form-item label="Assign categories">
              <el-select v-model="ruleDraft.assignCategories" multiple filterable :disabled="outlookBusy" placeholder="不設定分類">
                <el-option v-for="category in categories" :key="category.name" :label="category.name" :value="category.name" />
              </el-select>
            </el-form-item>
            <el-form-item label="Task interval">
              <el-select v-model="ruleDraft.markAsTaskInterval" :disabled="outlookBusy || !ruleDraft.markAsTask">
                <el-option label="Today" value="today" />
                <el-option label="Tomorrow" value="tomorrow" />
                <el-option label="This week" value="this_week" />
                <el-option label="Next week" value="next_week" />
                <el-option label="No date" value="no_date" />
              </el-select>
            </el-form-item>
          </div>
          <div class="rule-option-row">
            <el-checkbox v-model="ruleDraft.markAsTask" :disabled="outlookBusy">Mark as task</el-checkbox>
            <el-checkbox v-model="ruleDraft.clearCategories" :disabled="outlookBusy">Clear categories</el-checkbox>
            <el-checkbox v-model="ruleDraft.desktopAlert" :disabled="outlookBusy">Desktop alert</el-checkbox>
            <el-checkbox v-model="ruleDraft.delete" :disabled="outlookBusy">Move to Deleted Items</el-checkbox>
            <el-checkbox v-model="ruleDraft.stopProcessingMoreRules" :disabled="outlookBusy">Stop processing more rules</el-checkbox>
          </div>
          <div v-if="!ruleHasAction" class="rule-validation-hint">
            請至少選擇一個動作。
          </div>
        </section>
      </el-form>

      <template #footer>
        <div class="rule-editor-actions">
          <el-button :disabled="outlookBusy" @click="resetRuleDraft()">清除</el-button>
          <el-button type="primary" :icon="Select" :loading="operationLoading" :disabled="!ruleCanSave" @click="submitRule">
            儲存
          </el-button>
        </div>
      </template>
    </el-dialog>
  </main>
</template>
