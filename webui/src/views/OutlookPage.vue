<script setup lang="ts">
import { Document } from '@element-plus/icons-vue'
import CategoryManagerDialog from '../features/outlook/components/CategoryManagerDialog.vue'
import FlagEditorDialog from '../features/outlook/components/FlagEditorDialog.vue'
import FolderContextMenu from '../features/outlook/components/FolderContextMenu.vue'
import MailDetailDialog from '../features/outlook/components/MailDetailDialog.vue'
import MeetingInvitationDialog from '../features/outlook/components/MeetingInvitationDialog.vue'
import OutlookCalendarView from '../features/outlook/components/OutlookCalendarView.vue'
import OutlookChatView from '../features/outlook/components/OutlookChatView.vue'
import OutlookContactsView from '../features/outlook/components/OutlookContactsView.vue'
import OutlookHomeView from '../features/outlook/components/OutlookHomeView.vue'
import OutlookRulesView from '../features/outlook/components/OutlookRulesView.vue'
import OutlookSearchView from '../features/outlook/components/OutlookSearchView.vue'
import type { OutlookDashboardState } from '../features/outlook/composables/useOutlookDashboard'
import type { AppView } from '../features/outlook/models/outlook'
import { isMeetingMessage } from '../features/outlook/utils/outlookItemTypes'

const props = defineProps<{
  dashboard: OutlookDashboardState
}>()

const {
  activeView,
  addCategoryToMasterList,
  categories,
  categoryColorOptions,
  categoryColorStyle,
  categoryCreateColor,
  categoryCreateDraft,
  categoryManagerVisible,
  dialogMail,
  flagEditorVisible,
  hiddenMasterCategoryCount,
  loadingCategories,
  mailPropertiesDraft,
  masterCategoryListExpanded,
  navOptions,
  operationLoading,
  outlookBusy,
  requestCategories,
  signalRState,
  switchView,
  toggleMasterCategoryList,
  updateCategoryColor,
  visibleMasterCategories,
} = props.dashboard
</script>

<template>
  <div class="outlook-page">
    <div class="feature-toolbar">
      <div class="feature-title">
        <el-icon><Document /></el-icon>
        <span>Outlook</span>
        <el-tag :type="signalRState === 'connected' ? 'success' : 'danger'" effect="plain">
          {{ signalRState }}
        </el-tag>
      </div>

      <el-segmented
        :model-value="activeView"
        :options="navOptions"
        @update:model-value="(value: string | number | boolean) => switchView(value as AppView)"
      />
    </div>

    <OutlookHomeView v-if="activeView === 'outlook'" :dashboard="dashboard" />
    <OutlookSearchView v-else-if="activeView === 'search'" :dashboard="dashboard" />
    <OutlookRulesView v-else-if="activeView === 'rules'" :dashboard="dashboard" />
    <OutlookChatView v-else-if="activeView === 'chat'" :dashboard="dashboard" />
    <OutlookCalendarView v-else-if="activeView === 'calendar'" :dashboard="dashboard" />
    <OutlookContactsView v-else-if="activeView === 'contacts'" />

    <MeetingInvitationDialog v-if="dialogMail && isMeetingMessage(dialogMail)" :dashboard="dashboard" />
    <MailDetailDialog v-else :dashboard="dashboard" />
    <FolderContextMenu :dashboard="dashboard" />

    <CategoryManagerDialog
      v-model="categoryManagerVisible"
      v-model:category-create-color="categoryCreateColor"
      v-model:category-create-draft="categoryCreateDraft"
      :categories="categories"
      :category-color-options="categoryColorOptions"
      :category-color-style="categoryColorStyle"
      :hidden-master-category-count="hiddenMasterCategoryCount"
      :loading-categories="loadingCategories"
      :master-category-list-expanded="masterCategoryListExpanded"
      :operation-loading="operationLoading"
      :outlook-busy="outlookBusy"
      :visible-master-categories="visibleMasterCategories"
      @add-category="addCategoryToMasterList"
      @request-categories="requestCategories"
      @toggle-master-category-list="toggleMasterCategoryList"
      @update-category-color="updateCategoryColor"
    />

    <FlagEditorDialog
      v-model="flagEditorVisible"
      v-model:draft="mailPropertiesDraft"
      :outlook-busy="outlookBusy"
    />
  </div>
</template>
