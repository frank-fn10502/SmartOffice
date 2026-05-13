<script setup lang="ts">
import { ArrowRight, Folder, Search } from '@element-plus/icons-vue'
import OutlookFolderPane from './OutlookFolderPane.vue'
import SearchResultMailRow from './SearchResultMailRow.vue'
import type { OutlookDashboardState } from '../../composables/useOutlookDashboard'

const props = defineProps<{
  dashboard: OutlookDashboardState
}>()

const {
  categories,
  categoryTagStyle,
  clearMailDrag,
  flagDisplayLabel,
  flagTagType,
  loadingMailSearch,
  mailSearchDraft,
  mailSearchProgressText,
  mailSearchSummaryItems,
  mails,
  openMailDialog,
  outlookBusy,
  requestMailSearch,
  searchResultGroups,
  searchResultRows,
  searchResultViewMode,
  selectMail,
  selectedMailIds,
  splitCategories,
  startMailDrag,
  toggleSearchResultFolder,
  toggleSearchResultStore,
} = props.dashboard
</script>

<template>
  <main class="search-layout">
    <OutlookFolderPane :dashboard="dashboard" />

    <section class="panel search-page-panel">
      <div class="panel-header">
        <div class="panel-title">
          <el-icon><Search /></el-icon>
          <span>搜尋郵件</span>
          <el-tag effect="plain">{{ mails.length }}</el-tag>
        </div>
        <el-button type="primary" :loading="loadingMailSearch" :disabled="loadingMailSearch" @click="requestMailSearch">
          搜尋
        </el-button>
      </div>

      <div class="mail-search-panel standalone-search-panel">
        <div class="mail-search-flow">
          使用 Outlook 內建搜尋尋找符合文字與篩選條件的郵件；不做錯字模糊比對，也不限制結果筆數。
        </div>
        <div class="mail-search-row">
          <el-input
            v-model="mailSearchDraft.keyword"
            clearable
            :prefix-icon="Search"
            placeholder="片段關鍵字，例如：客戶xxxx"
            @keydown.enter.prevent="requestMailSearch"
          />
          <el-select v-model="mailSearchDraft.scopeMode" class="scope-select">
            <el-option label="目前資料夾" value="selected_folder" />
            <el-option label="目前信箱" value="selected_store" />
            <el-option label="全部信箱" value="global" />
          </el-select>
        </div>
        <div class="mail-search-row search-options-row">
          <span class="search-options-label">文字範圍</span>
          <el-checkbox-group v-model="mailSearchDraft.textFields">
            <el-checkbox label="subject">標題</el-checkbox>
            <el-checkbox label="sender">寄件者</el-checkbox>
            <el-checkbox label="body">內容</el-checkbox>
          </el-checkbox-group>
        </div>
        <div class="mail-search-row search-options-row">
          <span class="search-options-label">篩選條件</span>
          <el-select
            v-model="mailSearchDraft.categoryNames"
            class="category-filter-select"
            multiple
            collapse-tags
            collapse-tags-tooltip
            clearable
            placeholder="分類"
          >
            <el-option
              v-for="category in categories"
              :key="category.name"
              :label="category.name"
              :value="category.name"
            />
          </el-select>
          <el-select v-model="mailSearchDraft.hasAttachments" class="filter-select" clearable placeholder="附件">
            <el-option label="包含附件" :value="true" />
            <el-option label="不含附件" :value="false" />
          </el-select>
          <el-select v-model="mailSearchDraft.flagState" class="filter-select">
            <el-option label="旗標不限" value="any" />
            <el-option label="有旗標" value="flagged" />
            <el-option label="無旗標" value="unflagged" />
          </el-select>
          <el-select v-model="mailSearchDraft.readState" class="filter-select">
            <el-option label="已讀不限" value="any" />
            <el-option label="未讀" value="unread" />
            <el-option label="已讀" value="read" />
          </el-select>
        </div>
        <div class="mail-search-row search-options-row">
          <el-date-picker
            v-model="mailSearchDraft.receivedFrom"
            type="datetime"
            value-format="YYYY-MM-DDTHH:mm:ss"
            placeholder="收到時間起"
          />
          <el-date-picker
            v-model="mailSearchDraft.receivedTo"
            type="datetime"
            value-format="YYYY-MM-DDTHH:mm:ss"
            placeholder="收到時間迄"
          />
        </div>
      </div>

      <div class="search-result-criteria" :class="{ empty: mailSearchSummaryItems.length === 0 }">
        <div class="search-result-summary" :class="{ empty: mailSearchSummaryItems.length === 0 }">
          <template v-if="mailSearchSummaryItems.length > 0">
            <span
              v-for="item in mailSearchSummaryItems"
              :key="`${item.label}-${item.value}`"
              class="search-summary-chip"
              :class="item.tone"
            >
              <span>{{ item.label }}</span>
              <strong>{{ item.value }}</strong>
            </span>
          </template>
          <span v-else>尚未送出搜尋條件</span>
        </div>
      </div>

      <div class="search-result-toolbar">
        <div class="search-result-total">
          搜尋結果 <strong>{{ mails.length }}</strong> 封
        </div>
        <el-segmented
          v-model="searchResultViewMode"
          :options="[
            { label: 'Tree', value: 'tree' },
            { label: 'Flat', value: 'flat' },
          ]"
        />
      </div>
      <div v-if="loadingMailSearch" class="search-result-loading" role="status">
        <span>{{ mailSearchProgressText || 'Outlook 郵件搜尋中...' }}</span>
      </div>

      <div class="mail-table search-result-table">
        <p v-if="mails.length === 0 && !loadingMailSearch" class="hint">尚未取得搜尋結果。</p>
        <SearchResultMailRow
          v-for="{ mail, index, sourceLabel } in searchResultViewMode === 'flat' ? searchResultRows : []"
          :key="mail.id || `${mail.receivedTime}-${index}`"
          :mail="mail"
          :index="index"
          :source-label="sourceLabel"
          :selected-mail-ids="selectedMailIds"
          :outlook-busy="outlookBusy"
          :category-tag-style="categoryTagStyle"
          :flag-display-label="flagDisplayLabel"
          :flag-tag-type="flagTagType"
          :split-categories="splitCategories"
          @clear-mail-drag="clearMailDrag"
          @open-mail-dialog="openMailDialog"
          @select-mail="selectMail"
          @start-mail-drag="startMailDrag"
        />
        <div v-for="store in searchResultViewMode === 'tree' ? searchResultGroups : []" :key="store.key" class="search-result-store">
          <button class="search-result-tree-node store" type="button" @click="toggleSearchResultStore(store.key)">
            <el-icon class="search-result-disclosure" :class="{ collapsed: store.collapsed }"><ArrowRight /></el-icon>
            <el-icon><Folder /></el-icon>
            <span class="search-result-node-label">{{ store.label }}</span>
            <span class="search-result-count">{{ store.count }}</span>
          </button>
          <div v-if="!store.collapsed" class="search-result-store-children">
            <div v-for="folderGroup in store.folders" :key="folderGroup.key" class="search-result-folder">
              <button class="search-result-tree-node folder" type="button" @click="toggleSearchResultFolder(folderGroup.key)">
                <el-icon class="search-result-disclosure" :class="{ collapsed: folderGroup.collapsed }"><ArrowRight /></el-icon>
                <el-icon><Folder /></el-icon>
                <span class="search-result-node-label">{{ folderGroup.label }}</span>
                <span v-if="folderGroup.path" class="search-result-node-path">{{ folderGroup.path }}</span>
                <span class="search-result-count">{{ folderGroup.count }}</span>
              </button>
              <div v-if="!folderGroup.collapsed" class="search-result-folder-children">
                <SearchResultMailRow
                  v-for="{ mail, index } in folderGroup.rows"
                  :key="mail.id || `${mail.receivedTime}-${index}`"
                  class="search-result-tree-row"
                  :mail="mail"
                  :index="index"
                  :selected-mail-ids="selectedMailIds"
                  :outlook-busy="outlookBusy"
                  :category-tag-style="categoryTagStyle"
                  :flag-display-label="flagDisplayLabel"
                  :flag-tag-type="flagTagType"
                  :split-categories="splitCategories"
                  @clear-mail-drag="clearMailDrag"
                  @open-mail-dialog="openMailDialog"
                  @select-mail="selectMail"
                  @start-mail-drag="startMailDrag"
                />
              </div>
            </div>
          </div>
        </div>
      </div>
    </section>
  </main>
</template>
