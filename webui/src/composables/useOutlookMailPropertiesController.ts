import { computed, ref, watch } from 'vue'
import type { ComputedRef, Ref } from 'vue'
import type {
  MailItemDto,
  MailPropertiesCommandRequest,
  MailPropertiesDraft,
  OutlookCategoryDto,
} from '../models/outlook'
import {
  categoryOptionColor,
  categoryTextColor,
} from '../utils/categoryColors'
import {
  defaultFlagRequest,
  isDefaultFlagRequest,
  todayInputValue,
} from '../utils/outlookDashboardHelpers'
import { canUpdateMailProperties } from '../utils/outlookItemTypes'
import { outlookApi } from '../api/outlook'
import { buildMailPropertiesDraft, buildMailPropertiesPayload } from './outlookMailProperties'

type MailPropertiesControllerOptions = {
  activeMailForProperties: ComputedRef<MailItemDto | null>
  categories: Ref<OutlookCategoryDto[]>
  loadCachedCategories: () => Promise<void>
  loadCachedMailSearchResults: () => Promise<void>
  loadCachedMails: () => Promise<void>
  outlookBusy: ComputedRef<boolean>
  runMailOperation: (action: () => Promise<unknown>, afterSuccess?: () => Promise<void>) => Promise<boolean>
  upsertCategory: (name: string, color: string, shortcutKey?: string) => Promise<unknown>
}

export function useOutlookMailPropertiesController(options: MailPropertiesControllerOptions) {
  const {
    activeMailForProperties,
    categories,
    loadCachedCategories,
    loadCachedMailSearchResults,
    loadCachedMails,
    outlookBusy,
    runMailOperation,
    upsertCategory,
  } = options

  const activeMailPropertySections = ref(['set-mail-properties'])
  const categoryManagerVisible = ref(false)
  const flagEditorVisible = ref(false)
  const masterCategoryListExpanded = ref(false)
  const categoryCreateDraft = ref('')
  const categoryCreateColor = ref('olCategoryColorNone')
  const mailPropertiesDraft = ref<MailPropertiesDraft>({
    isRead: false,
    flagInterval: 'none',
    flagRequest: '',
    taskStartDate: '',
    taskDueDate: '',
    taskCompletedDate: '',
    categories: [] as string[],
  })

  const visibleMasterCategories = computed(() => (
    masterCategoryListExpanded.value ? categories.value : categories.value.slice(0, 5)
  ))

  const hiddenMasterCategoryCount = computed(() => Math.max(0, categories.value.length - visibleMasterCategories.value.length))

  const mailPropertiesChanged = computed(() => {
    if (!activeMailForProperties.value) return false
    return JSON.stringify(buildMailPropertiesPayload(activeMailForProperties.value, buildMailPropertiesDraft(activeMailForProperties.value)))
      !== JSON.stringify(buildMailPropertiesPayload(activeMailForProperties.value, mailPropertiesDraft.value))
  })

  function categoryColorByName(name: string) {
    const category = categories.value.find((item) => item.name.toLowerCase() === name.trim().toLowerCase())
    return categoryOptionColor(category?.color)
  }

  function categoryTagStyle(name: string) {
    const color = categoryColorByName(name)
    return {
      '--el-tag-border-color': color,
      '--el-tag-bg-color': color,
      '--el-tag-text-color': categoryTextColor(color),
      backgroundColor: color,
      borderColor: color,
      color: categoryTextColor(color),
    }
  }

  function resetMailPropertiesDraft(mail: MailItemDto | null) {
    if (!mail) return
    mailPropertiesDraft.value = buildMailPropertiesDraft(mail)
  }

  async function applyMailProperties(mail: MailItemDto) {
    if (!mail.id?.trim() || !canUpdateMailProperties(mail)) return
    const payload = buildMailPropertiesPayload(mail, mailPropertiesDraft.value)
    const existingCategoryNames = new Set(categories.value.map((category) => category.name.toLowerCase()))
    const newCategories = payload.categories
      .filter((category) => !existingCategoryNames.has(category.toLowerCase()))
      .map((name) => ({ name, color: 'olCategoryColorNone', colorValue: 0, shortcutKey: '' }))
    const body: MailPropertiesCommandRequest = {
      ...payload,
      newCategories,
    }
    await runMailOperation(
      () => outlookApi.requestUpdateMailProperties(body),
      async () => {
        await Promise.allSettled([loadCachedMails(), loadCachedMailSearchResults(), loadCachedCategories()])
      },
    )
  }

  function setMailFlagDraft(interval: string) {
    if (outlookBusy.value) return
    mailPropertiesDraft.value.flagInterval = interval
    if (interval === 'custom') openFlagEditor()
  }

  function addMailCategoryDraft(categoryName: string) {
    const name = categoryName.trim()
    if (!name || outlookBusy.value) return
    const selected = mailPropertiesDraft.value.categories
    const exists = selected.some((category) => category.toLowerCase() === name.toLowerCase())
    if (!exists) mailPropertiesDraft.value.categories = [...selected, name]
  }

  function removeMailCategoryDraft(categoryName: string) {
    if (outlookBusy.value) return
    const name = categoryName.trim().toLowerCase()
    mailPropertiesDraft.value.categories = mailPropertiesDraft.value.categories
      .filter((category) => category.toLowerCase() !== name)
  }

  function toggleMasterCategoryList() {
    masterCategoryListExpanded.value = !masterCategoryListExpanded.value
  }

  function openCategoryManager() {
    categoryManagerVisible.value = true
  }

  function openFlagEditor() {
    if (outlookBusy.value || mailPropertiesDraft.value.flagInterval !== 'custom') return
    flagEditorVisible.value = true
  }

  async function addCategoryToMasterList() {
    const name = categoryCreateDraft.value.trim()
    if (!name) return
    await upsertCategory(name, categoryCreateColor.value)
    categoryCreateDraft.value = ''
    categoryCreateColor.value = 'olCategoryColorNone'
  }

  async function updateCategoryColor(category: OutlookCategoryDto, color: string) {
    await upsertCategory(category.name, color, category.shortcutKey)
  }

  watch(
    () => activeMailForProperties.value?.id,
    () => resetMailPropertiesDraft(activeMailForProperties.value),
  )

  watch(
    () => mailPropertiesDraft.value.flagInterval,
    (next, previous) => {
      if (isDefaultFlagRequest(mailPropertiesDraft.value.flagRequest, previous)) {
        mailPropertiesDraft.value.flagRequest = defaultFlagRequest(next)
      }
      if (next === 'custom') {
        const today = todayInputValue()
        mailPropertiesDraft.value.taskStartDate ||= today
        mailPropertiesDraft.value.taskDueDate ||= today
      } else {
        flagEditorVisible.value = false
      }
    },
  )

  return {
    activeMailPropertySections,
    addCategoryToMasterList,
    addMailCategoryDraft,
    applyMailProperties,
    categoryCreateColor,
    categoryCreateDraft,
    categoryManagerVisible,
    categoryTagStyle,
    flagEditorVisible,
    hiddenMasterCategoryCount,
    mailPropertiesChanged,
    mailPropertiesDraft,
    masterCategoryListExpanded,
    openCategoryManager,
    removeMailCategoryDraft,
    resetMailPropertiesDraft,
    setMailFlagDraft,
    toggleMasterCategoryList,
    updateCategoryColor,
    visibleMasterCategories,
  }
}
