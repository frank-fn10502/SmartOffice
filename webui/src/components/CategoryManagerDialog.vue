<script setup lang="ts">
import { Refresh } from '@element-plus/icons-vue'
import type { OutlookCategoryDto } from '../models/outlook'

const visible = defineModel<boolean>({ required: true })
const categoryCreateDraft = defineModel<string>('categoryCreateDraft', { required: true })
const categoryCreateColor = defineModel<string>('categoryCreateColor', { required: true })

defineProps<{
  categories: OutlookCategoryDto[]
  categoryColorOptions: { label: string; value: string }[]
  hiddenMasterCategoryCount: number
  loadingCategories: boolean
  masterCategoryListExpanded: boolean
  operationLoading: boolean
  outlookBusy: boolean
  visibleMasterCategories: OutlookCategoryDto[]
  categoryColorStyle: (value?: string) => Record<string, string>
}>()

defineEmits<{
  addCategory: []
  requestCategories: []
  toggleMasterCategoryList: []
  updateCategoryColor: [category: OutlookCategoryDto, color: string]
}>()
</script>

<template>
  <el-dialog v-model="visible" title="管理分類" width="520px" append-to-body>
    <div class="dialog-form">
      <div class="category-heading-row">
        <div class="library-heading">Master Category List</div>
        <el-button
          :icon="Refresh"
          circle
          size="small"
          :loading="loadingCategories"
          :disabled="outlookBusy && !loadingCategories"
          @click="$emit('requestCategories')"
        />
        <el-button size="small" text :disabled="categories.length <= 5" @click="$emit('toggleMasterCategoryList')">
          {{ masterCategoryListExpanded ? '縮回' : `全部展開${hiddenMasterCategoryCount ? ` (${hiddenMasterCategoryCount})` : ''}` }}
        </el-button>
      </div>
      <div class="category-add-row">
        <el-input
          v-model="categoryCreateDraft"
          :disabled="outlookBusy"
          placeholder="新增或更新分類名稱"
          @keydown.enter.prevent="$emit('addCategory')"
        />
        <el-select v-model="categoryCreateColor" class="category-color-select" :disabled="outlookBusy">
          <el-option
            v-for="option in categoryColorOptions"
            :key="option.value"
            :label="option.label"
            :value="option.value"
          >
            <span class="category-option">
              <span class="category-swatch" :style="categoryColorStyle(option.value)" />
              <span>{{ option.label }}</span>
            </span>
          </el-option>
        </el-select>
        <el-button :loading="operationLoading" :disabled="outlookBusy || !categoryCreateDraft.trim()" @click="$emit('addCategory')">
          儲存
        </el-button>
      </div>
      <div class="category-list compact-category-list">
        <div v-if="categories.length === 0 && !loadingCategories" class="hint">尚未取得 Outlook master category list。</div>
        <div v-for="category in visibleMasterCategories" :key="category.name" class="category-row">
          <span class="category-name">
            <span class="category-swatch" :style="categoryColorStyle(category.color)" />
            <span>{{ category.name }}</span>
          </span>
          <el-select
            :model-value="category.color || 'olCategoryColorNone'"
            class="category-row-select"
            :disabled="outlookBusy"
            @change="(value: string | number | boolean) => $emit('updateCategoryColor', category, String(value))"
          >
            <el-option
              v-for="option in categoryColorOptions"
              :key="option.value"
              :label="option.label"
              :value="option.value"
            >
              <span class="category-option">
                <span class="category-swatch" :style="categoryColorStyle(option.value)" />
                <span>{{ option.label }}</span>
              </span>
            </el-option>
          </el-select>
        </div>
        <div v-if="loadingCategories" class="pane-loading">
          <span>Outlook category 同步中...</span>
        </div>
      </div>
    </div>
  </el-dialog>
</template>

<style scoped>
.dialog-form {
  display: grid;
  gap: 12px;
  min-width: 0;
}

.library-heading {
  color: #667085;
  font-size: 0.76rem;
  font-weight: 800;
  text-transform: uppercase;
}

.category-heading-row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  min-width: 0;
}

.category-add-row {
  display: grid;
  grid-template-columns: minmax(0, 1fr) 118px auto;
  gap: 8px;
}

.category-list {
  display: grid;
  gap: 8px;
  position: relative;
}

.compact-category-list {
  max-height: 260px;
  overflow: auto;
  min-height: 0;
}

.category-row {
  display: grid;
  grid-template-columns: minmax(0, 1fr) 118px;
  align-items: center;
  gap: 8px;
  padding: 8px;
  border: 1px solid #edf1f5;
  border-radius: 6px;
  background: #fff;
}

.category-name,
.category-option {
  display: inline-flex;
  align-items: center;
  min-width: 0;
  gap: 8px;
}

.category-name span:last-child {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.category-swatch {
  width: 14px;
  height: 14px;
  flex: 0 0 auto;
  border: 1px solid rgb(15 23 42 / 16%);
  border-radius: 50%;
}

.category-color-select,
.category-row-select {
  width: 100%;
}

@media (max-width: 680px) {
  .category-add-row,
  .category-row {
    grid-template-columns: 1fr;
  }
}
</style>
