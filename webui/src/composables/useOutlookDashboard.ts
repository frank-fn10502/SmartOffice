import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import { ElMessage } from 'element-plus'
import * as signalR from '@microsoft/signalr'
import {
  normalizeMailAttachments,
  normalizeMailItems,
  normalizeMailSearchProgress,
  normalizeOutlookCategories,
  outlookApi,
} from '../api/outlook'
import type {
  AddinLogEntry,
  AddinStatusDto,
  AttachmentExportSettingsDto,
  AppView,
  CalendarEventDto,
  ChatMessageDto,
  FolderTreeNode,
  MailAttachmentDto,
  MailAttachmentsDto,
  MailConversationDto,
  MailItemDto,
  MailPropertiesDraft,
  MailSearchProgressDto,
  MailPropertiesCommandRequest,
  OutlookStoreDto,
  OutlookCategoryDto,
  OutlookRuleDto,
  OutlookRuleCommandRequest,
  SignalRState,
} from '../models/outlook'
import {
  categoryColorOptions,
  categoryColorStyle,
  categoryColorValue,
  categoryOptionColor,
  categoryTextColor,
} from '../utils/categoryColors'
import { buildFolderTree, collectFolderOptions, findFolderByPath, folderType, isMailSelectableFolder, visibleRootFolders } from '../utils/folders'
import {
  addMonths,
  buildCalendarWeeks,
  dateInputToIso,
  defaultFlagRequest,
  flagDisplayLabel,
  flagIntervalOptions,
  flagTagType,
  isDefaultFlagRequest,
  monthEndExclusive,
  monthStart,
  splitCategories,
  toDateInput,
  toDateKey,
  todayInputValue,
} from '../utils/outlookDashboardHelpers'

function estimatedAttachmentExportRoot() {
  const platform = window.navigator.platform.toLowerCase()
  if (platform.includes('win')) return 'E:\\SmartOffice\\Attachments、D:\\SmartOffice\\Attachments 或 C:\\SmartOffice\\Attachments'
  return '$HOME/SmartOffice/Attachments'
}

const manualOutlookDeleteMessage = 'SmartOffice API 不會永久刪除 Outlook 郵件或 folder。此項目已在 Outlook 刪除資料夾內；若要永久刪除，請到 Outlook 手動操作。'
const mailFetchDelayMs = 300
const mailFetchCountdownTickMs = 100

type RuleDraft = {
  storeId: string
  ruleName: string
  originalRuleName: string
  originalExecutionOrder?: number
  ruleType: 'receive' | 'send'
  enabled: boolean
  subjectContains: string
  bodyContains: string
  senderAddressContains: string
  categories: string[]
  hasAttachment: 'any' | 'yes'
  moveToFolderPath: string
  assignCategories: string[]
  markAsTask: boolean
  stopProcessingMoreRules: boolean
}

export function useOutlookDashboard() {
  const activeView = ref<AppView>('outlook')
  const signalRState = ref<SignalRState>('disconnected')
  const folders = ref<FolderTreeNode[]>([])
  const folderStores = ref<OutlookStoreDto[]>([])
  const folderMails = ref<MailItemDto[]>([])
  const mailSearchResults = ref<MailItemDto[]>([])
  const mailListMode = ref<'folder' | 'search'>('folder')
  const rules = ref<OutlookRuleDto[]>([])
  const selectedRuleIndex = ref<number | null>(null)
  const ruleDraft = ref<RuleDraft>({
    storeId: '',
    ruleName: '',
    originalRuleName: '',
    originalExecutionOrder: undefined as number | undefined,
    ruleType: 'receive' as 'receive' | 'send',
    enabled: true,
    subjectContains: '',
    bodyContains: '',
    senderAddressContains: '',
    categories: [] as string[],
    hasAttachment: 'any' as 'any' | 'yes',
    moveToFolderPath: '',
    assignCategories: [] as string[],
    markAsTask: false,
    stopProcessingMoreRules: true,
  })
  const categories = ref<OutlookCategoryDto[]>([])
  const calendarEvents = ref<CalendarEventDto[]>([])
  const calendarMonthDate = ref(monthStart(new Date()))
  const selectedCalendarEvent = ref<CalendarEventDto | null>(null)
  const chatMessages = ref<ChatMessageDto[]>([])
  const addinStatus = ref<AddinStatusDto>({
    connected: false,
    lastCommand: '',
  })
  const addinLogs = ref<AddinLogEntry[]>([])
  const estimatedExportRoot = estimatedAttachmentExportRoot()
  const attachmentExportSettings = ref<AttachmentExportSettingsDto>({
    rootPath: estimatedExportRoot,
    defaultRootPath: estimatedExportRoot,
  })
  const attachmentExportRootDraft = ref(estimatedExportRoot)
  const savingAttachmentExportSettings = ref(false)
  const selectedFolderPath = ref('')
  const fetchedMailFolderPath = ref('')
  const pendingMailFolderPath = ref('')
  const selectedMailIndex = ref<number | null>(null)
  const selectedMailIds = ref<Set<string>>(new Set())
  const mailDialogVisible = ref(false)
  const mailDialogIndex = ref<number | null>(null)
  const mailDialogMailId = ref('')
  const mailDialogHtml = ref(false)
  const activeMailPropertySections = ref(['set-mail-properties'])
  const expandedFolders = ref<Set<string>>(new Set())
  const loadingMailBodyIds = ref<Set<string>>(new Set())
  const mailAttachmentsByMailId = ref<Record<string, MailAttachmentDto[]>>({})
  const mailConversationsByMailId = ref<Record<string, MailConversationDto>>({})
  const loadingAttachmentMailIds = ref<Set<string>>(new Set())
  const loadingConversationMailIds = ref<Set<string>>(new Set())
  const exportingAttachmentIds = ref<Set<string>>(new Set())
  const mailLookbackHours = ref(168)
  const mailCount = ref(30)
  const lastMailFetchAt = ref<Date | null>(null)
  const scheduledMailFetchAt = ref(0)
  const mailFetchCountdownTick = ref(Date.now())
  const loadingMailSearch = ref(false)
  const searchResultViewMode = ref<'flat' | 'tree'>('tree')
  const collapsedSearchResultStores = ref<Set<string>>(new Set())
  const collapsedSearchResultFolders = ref<Set<string>>(new Set())
  const mailSearchProgress = ref<MailSearchProgressDto | null>(null)
  const mailSearchDraft = ref({
    keyword: '',
    textFields: ['subject'] as Array<'subject' | 'sender' | 'body'>,
    categoryNames: [] as string[],
    hasAttachments: undefined as boolean | undefined,
    flagState: 'any' as 'any' | 'flagged' | 'unflagged',
    readState: 'any' as 'any' | 'unread' | 'read',
    receivedFrom: '',
    receivedTo: '',
    scopeMode: 'selected_folder' as 'selected_folder' | 'selected_store' | 'global',
  })
  const activeMailSearchSummary = ref<Array<{ label: string; value: string; tone: 'active' | 'muted' | 'info' }>>([])
  const chatText = ref('')
  const loadingFolders = ref(false)
  const loadingMails = ref(true)
  const loadingRules = ref(false)
  const loadingCategories = ref(true)
  const loadingCalendar = ref(false)
  const loadingSignalRPing = ref(false)
  const requestLoading = ref(false)
  const outlookFirstLoadCompleted = ref(false)
  const mailPropertiesDraft = ref<MailPropertiesDraft>({
    isRead: false,
    flagInterval: 'none',
    flagRequest: '',
    taskStartDate: '',
    taskDueDate: '',
    taskCompletedDate: '',
    categories: [] as string[],
  })
  const categoryManagerVisible = ref(false)
  const flagEditorVisible = ref(false)
  const masterCategoryListExpanded = ref(false)
  const categoryCreateDraft = ref('')
  const categoryCreateColor = ref('olCategoryColorNone')
  const creatingFolderParentPath = ref('')
  const creatingFolderName = ref('')
  const draggedMailId = ref('')
  const dragOverFolderPath = ref('')
  const folderContextMenu = ref({
    visible: false,
    x: 0,
    y: 0,
    folderPath: '',
  })
  const chatPanelRef = ref<HTMLElement | null>(null)
  const mailHtmlSandbox = 'allow-same-origin allow-popups allow-popups-to-escape-sandbox'
  let connection: signalR.HubConnection | null = null
  let unmounted = false
  let initialFoldersFetchCompleted = false
  let initialMailsFetchCompleted = false
  let initialCategoriesFetchCompleted = false
  let startupSyncStarted = false
  let activeRequestId = ''
  let requestTimeoutId = 0
  let mailFetchTimeoutId = 0
  let mailFetchCountdownIntervalId = 0
  let lastSelectedMailIndex = -1
  const mailBodyTimeoutIds = new Map<string, number>()
  const attachmentTimeoutIds = new Map<string, number>()

  const visibleFolders = computed(() => visibleRootFolders(folders.value))
  const mails = computed(() => mailListMode.value === 'search' ? mailSearchResults.value : folderMails.value)
  const searchResultRows = computed(() => mailSearchResults.value.map((mail, index) => ({
    mail,
    index,
    sourceLabel: mailSourceLabel(mail),
  })))
  const searchResultGroups = computed(() => {
    const groups = new Map<string, {
      key: string
      label: string
      count: number
      collapsed: boolean
      folders: {
        key: string
        label: string
        path: string
        count: number
        collapsed: boolean
        rows: typeof searchResultRows.value
      }[]
    }>()
    for (const row of searchResultRows.value) {
      const source = mailSource(row.mail)
      const storeKey = source.storeId || source.storeLabel
      const folderKey = `${storeKey}\n${source.folderPath || source.folderLabel}`
      let store = groups.get(storeKey)
      if (!store) {
        store = {
          key: storeKey,
          label: source.storeLabel,
          count: 0,
          collapsed: collapsedSearchResultStores.value.has(storeKey),
          folders: [],
        }
        groups.set(storeKey, store)
      }
      let folder = store.folders.find((item) => item.key === folderKey)
      if (!folder) {
        folder = {
          key: folderKey,
          label: source.folderLabel,
          path: source.folderPath,
          count: 0,
          collapsed: collapsedSearchResultFolders.value.has(folderKey),
          rows: [],
        }
        store.folders.push(folder)
      }
      store.count += 1
      folder.count += 1
      folder.rows.push(row)
    }
    return [...groups.values()]
  })

  const mailStats = computed(() => ({
    unread: mails.value.filter((mail) => !mail.isRead).length,
    flagged: mails.value.filter((mail) => mail.isMarkedAsTask).length,
    highImportance: mails.value.filter((mail) => mail.importance === 'high').length,
    categorized: mails.value.filter((mail) => Boolean(mail.categories)).length,
  }))
  const visibleMasterCategories = computed(() => (
    masterCategoryListExpanded.value ? categories.value : categories.value.slice(0, 5)
  ))
  const hiddenMasterCategoryCount = computed(() => Math.max(0, categories.value.length - visibleMasterCategories.value.length))

  const outlookBusy = computed(() => {
    return loadingFolders.value || loadingMails.value || loadingRules.value || loadingCategories.value || loadingCalendar.value || requestLoading.value
  })

  const outlookBusyText = computed(() => {
    if (loadingFolders.value) return 'Outlook folder 同步中...'
    if (loadingMails.value) return 'Outlook 郵件抓取中...'
    if (loadingRules.value) return 'Outlook rule 同步中...'
    if (loadingCategories.value) return 'Outlook category 同步中...'
    if (loadingCalendar.value) return 'Outlook calendar 同步中...'
    if (requestLoading.value) return 'Outlook 操作執行中...'
    return ''
  })

  const outlookDependentViewsLocked = computed(() => !outlookFirstLoadCompleted.value)
  const navOptions = computed(() => [
    { label: 'Outlook', value: 'outlook' },
    { label: 'Search', value: 'search', disabled: outlookDependentViewsLocked.value },
    { label: 'Rules', value: 'rules', disabled: outlookDependentViewsLocked.value },
    { label: 'Chat', value: 'chat', disabled: outlookDependentViewsLocked.value },
    { label: 'Calendar', value: 'calendar', disabled: outlookDependentViewsLocked.value },
  ])

  const folderOptions = computed(() => collectFolderOptions(visibleFolders.value))

  const calendarWeekdays = ['日', '一', '二', '三', '四', '五', '六']

  const calendarMonthLabel = computed(() => {
    return calendarMonthDate.value.toLocaleDateString('zh-TW', { year: 'numeric', month: 'long' })
  })

  const calendarWeeks = computed(() => buildCalendarWeeks(calendarMonthDate.value, calendarEvents.value))

  const selectedFolderName = computed(() => {
    return folderNameForPath(selectedFolderPath.value)
  })

  const fetchedMailFolderName = computed(() => {
    if (mailListMode.value === 'search') return `搜尋結果：${mailSearchResults.value.length}`
    return fetchedMailFolderPath.value ? folderNameForPath(fetchedMailFolderPath.value) : '尚未抓取郵件'
  })

  const mailSearchProgressText = computed(() => {
    const progress = mailSearchProgress.value
    if (!progress || !loadingMailSearch.value) return ''
    const scopeText = progress.totalFolders > 0
      ? `${progress.processedFolders}/${progress.totalFolders} folders`
      : progress.totalStores > 0
        ? `${progress.processedStores}/${progress.totalStores} stores`
        : '準備中'
    const current = progress.currentFolderPath ? ` · ${folderNameForPath(progress.currentFolderPath)}` : ''
    return `${progress.percent}% · ${scopeText}${current}`
  })

  const mailSearchSummaryItems = computed(() => activeMailSearchSummary.value)

  const mailListNeedsFetch = computed(() => {
    if (mailListMode.value === 'search') return false
    return Boolean(selectedFolderPath.value && selectedFolderPath.value !== fetchedMailFolderPath.value)
  })

  const mailFetchCountdownSeconds = computed(() => {
    if (!scheduledMailFetchAt.value) return 0
    return Math.max(0, Math.ceil((scheduledMailFetchAt.value - mailFetchCountdownTick.value) / 100) / 10)
  })

  const mailFetchCountdownText = computed(() => {
    if (!scheduledMailFetchAt.value) return ''
    return `${mailFetchCountdownSeconds.value.toFixed(1)} 秒後自動抓取`
  })

  const showMailFetchWarning = computed(() => {
    return mailListNeedsFetch.value
      && !mailFetchCountdownText.value
      && !loadingMails.value
  })

  const mailFetchStatusText = computed(() => {
    if (mailFetchCountdownText.value) return `已選取 ${selectedFolderName.value}，${mailFetchCountdownText.value}；按「立即抓取」可提早執行。`
    if (loadingMails.value && pendingMailFolderPath.value) return `正在抓取：${folderNameForPath(pendingMailFolderPath.value)}`
    if (lastMailFetchAt.value) return `上次抓取：${lastMailFetchAt.value.toLocaleTimeString('zh-TW', { hour: '2-digit', minute: '2-digit', second: '2-digit' })}`
    return '尚未抓取郵件。'
  })

  const contextFolderName = computed(() => {
    return folderOptions.value.find((folder) => folder.folderPath === folderContextMenu.value.folderPath)?.label.trim() ?? '未選擇'
  })

  const selectedMail = computed(() => {
    if (selectedMailIndex.value === null) return null
    return mails.value[selectedMailIndex.value] ?? null
  })

  const selectedMailCategories = computed(() => {
    if (!selectedMail.value?.categories) return []
    return selectedMail.value.categories
      .split(',')
      .map((category) => category.trim())
      .filter(Boolean)
  })

  const dialogMail = computed(() => {
    if (mailDialogMailId.value) {
      const mail = mails.value.find((item) => item.id === mailDialogMailId.value)
        ?? folderMails.value.find((item) => item.id === mailDialogMailId.value)
        ?? mailSearchResults.value.find((item) => item.id === mailDialogMailId.value)
      if (mail) return mail
    }
    if (mailDialogIndex.value === null) return null
    return mails.value[mailDialogIndex.value] ?? null
  })

  const dialogMailAttachments = computed(() => {
    return dialogMail.value?.id ? mailAttachmentsByMailId.value[dialogMail.value.id] ?? [] : []
  })

  const dialogMailConversation = computed(() => {
    return dialogMail.value?.id ? mailConversationsByMailId.value[dialogMail.value.id] ?? null : null
  })

  const dialogMailConversationItems = computed(() => dialogMailConversation.value?.mails ?? [])

  const selectedRule = computed(() => {
    return selectedRuleIndex.value === null ? null : rules.value[selectedRuleIndex.value] ?? null
  })

  const ruleDraftIsEditing = computed(() => Boolean(ruleDraft.value.originalRuleName))

  const dialogLoading = computed(() => {
    const mail = dialogMail.value
    return Boolean(mail && (isMailBodyLoading(mail) || isAttachmentListLoading(mail) || isConversationLoading(mail)))
  })

  const dialogMailFolderName = computed(() => {
    return dialogMail.value?.folderPath ? folderNameForPath(dialogMail.value.folderPath) : '未選擇'
  })

  const dialogMailHasIdentity = computed(() => Boolean(dialogMail.value?.id?.trim()))

  const activeMailForProperties = computed(() => dialogMail.value ?? selectedMail.value)

  const mailPropertiesChanged = computed(() => {
    if (!activeMailForProperties.value) return false
    return JSON.stringify(buildMailPropertiesPayload(activeMailForProperties.value, buildMailPropertiesDraft(activeMailForProperties.value)))
      !== JSON.stringify(buildMailPropertiesPayload(activeMailForProperties.value, mailPropertiesDraft.value))
  })

  function folderNameForPath(path: string) {
    if (!path) return '未選擇'
    return folderOptions.value.find((folder) => folder.folderPath === path)?.label.trim() ?? path
  }

  function storeForFolderPath(path: string) {
    if (!path) return undefined
    return folderStores.value.find((store) => {
      const root = store.rootFolderPath
      return root && (
        path === root
        || path.startsWith(`${root}/`)
        || path.startsWith(`${root}\\`)
      )
    })
  }

  function folderLeafName(path: string) {
    const parts = path.split(/[\\/]+/).map((part) => part.trim()).filter(Boolean)
    return parts.at(-1) || path || 'Unknown folder'
  }

  function mailSource(mail: MailItemDto) {
    const folder = folderOptions.value.find((item) => item.folderPath === mail.folderPath)
    const store = folderStores.value.find((item) => item.storeId === folder?.storeId)
      ?? storeForFolderPath(mail.folderPath)
    const storeLabel = searchStoreLabel(store, folder?.storeId)
    const folderLabel = folder?.name || folderLeafName(mail.folderPath)
    return {
      storeId: store?.storeId || folder?.storeId || '',
      storeLabel,
      folderLabel,
      folderPath: mail.folderPath,
    }
  }

  function searchStoreLabel(store: OutlookStoreDto | undefined, fallbackStoreId = '') {
    if (!store) return fallbackStoreId || 'Unknown store'
    const kind = store.storeKind?.trim().toUpperCase() || 'STORE'
    if (kind === 'PST') {
      const fileName = store.storeFilePath.split(/[\\/]+/).filter(Boolean).at(-1)
      return `PST · ${fileName || store.displayName || store.storeId}`
    }
    if (kind === 'OST') return `OST · ${store.displayName || store.storeId}`
    return `${kind} · ${store.displayName || store.storeId}`
  }

  function mailSourceLabel(mail: MailItemDto) {
    const source = mailSource(mail)
    return source.folderPath ? `${source.storeLabel} / ${source.folderLabel}` : source.storeLabel
  }

  function compareMailSearchResults(left: MailItemDto, right: MailItemDto) {
    const leftLabel = mailSourceLabel(left)
    const rightLabel = mailSourceLabel(right)
    const sourceOrder = leftLabel.localeCompare(rightLabel, undefined, { sensitivity: 'base' })
    if (sourceOrder !== 0) return sourceOrder
    return new Date(right.receivedTime).getTime() - new Date(left.receivedTime).getTime()
  }

  function setMailSearchResults(items: MailItemDto[]) {
    mailSearchResults.value = [...items].sort(compareMailSearchResults)
    if (mailListMode.value === 'search') pruneSelectedMailIds(mailSearchResults.value)
  }

  function toggleSearchResultStore(key: string) {
    const next = new Set(collapsedSearchResultStores.value)
    if (next.has(key)) next.delete(key)
    else next.add(key)
    collapsedSearchResultStores.value = next
  }

  function toggleSearchResultFolder(key: string) {
    const next = new Set(collapsedSearchResultFolders.value)
    if (next.has(key)) next.delete(key)
    else next.add(key)
    collapsedSearchResultFolders.value = next
  }

  function inferMailFolderPath(items: MailItemDto[], fallback = '') {
    const paths = [...new Set(items.map((mail) => mail.folderPath).filter(Boolean))]
    return paths.length === 1 ? paths[0] : fallback
  }

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

  function buildMailPropertiesDraft(mail: MailItemDto) {
    const flagInterval = mail.flagInterval || (mail.isMarkedAsTask ? 'today' : 'none')
    return {
      isRead: mail.isRead,
      flagInterval,
      flagRequest: isDefaultFlagRequest(mail.flagRequest) ? defaultFlagRequest(flagInterval) : mail.flagRequest,
      taskStartDate: toDateInput(mail.taskStartDate),
      taskDueDate: toDateInput(mail.taskDueDate),
      taskCompletedDate: toDateInput(mail.taskCompletedDate),
      categories: splitCategories(mail.categories),
    }
  }

  function normalizeMailPropertiesDraft(draft: typeof mailPropertiesDraft.value) {
    return {
      isRead: draft.isRead,
      flagInterval: draft.flagInterval || 'none',
      flagRequest: draft.flagInterval === 'none' ? '' : (draft.flagRequest || defaultFlagRequest(draft.flagInterval)).trim(),
      taskStartDate: draft.taskStartDate || '',
      taskDueDate: draft.taskDueDate || '',
      taskCompletedDate: draft.taskCompletedDate || '',
      categories: [...new Set(draft.categories.map((category) => category.trim()).filter(Boolean))]
        .sort((left, right) => left.localeCompare(right, undefined, { sensitivity: 'base' })),
    }
  }

  function buildMailPropertiesPayload(mail: MailItemDto, draft: typeof mailPropertiesDraft.value) {
    const normalized = normalizeMailPropertiesDraft(draft)
    const isCustomFlag = normalized.flagInterval === 'custom'
    return {
      mailId: mail.id,
      folderPath: mail.folderPath,
      isRead: normalized.isRead,
      flagInterval: normalized.flagInterval,
      flagRequest: normalized.flagRequest,
      taskStartDate: isCustomFlag ? dateInputToIso(normalized.taskStartDate) : undefined,
      taskDueDate: isCustomFlag ? dateInputToIso(normalized.taskDueDate) : undefined,
      taskCompletedDate: normalized.flagInterval === 'complete' ? dateInputToIso(normalized.taskCompletedDate) : undefined,
      categories: normalized.categories,
    }
  }

  function resetMailPropertiesDraft(mail: MailItemDto | null) {
    if (!mail) return
    mailPropertiesDraft.value = buildMailPropertiesDraft(mail)
  }

  function updateOutlookFirstLoadCompleted() {
    outlookFirstLoadCompleted.value = initialFoldersFetchCompleted
      && initialMailsFetchCompleted
      && initialCategoriesFetchCompleted
  }

  function setMails(items: MailItemDto[], preferredMailId = selectedMail.value?.id ?? '') {
    const sortedItems = [...items].sort((left, right) =>
      new Date(right.receivedTime).getTime() - new Date(left.receivedTime).getTime()
    )
    folderMails.value = sortedItems
    mailListMode.value = 'folder'
    pruneSelectedMailIds(sortedItems)

    if (sortedItems.length === 0) {
      selectedMailIndex.value = null
      lastSelectedMailIndex = -1
      return
    }

    const nextIndex = preferredMailId ? sortedItems.findIndex((mail) => mail.id === preferredMailId) : -1
    selectedMailIndex.value = nextIndex >= 0 ? nextIndex : 0
  }

  function pruneSelectedMailIds(items = mails.value) {
    const visibleIds = new Set(items.map((mail) => mail.id).filter(Boolean))
    selectedMailIds.value = new Set([...selectedMailIds.value].filter((id) => visibleIds.has(id)))
    if (lastSelectedMailIndex >= items.length) lastSelectedMailIndex = -1
  }

  function selectedBulkMoveMails() {
    return mails.value.filter((mail) => mail.id && selectedMailIds.value.has(mail.id))
  }

  function selectedBulkMoveSourcePaths() {
    return new Set(selectedBulkMoveMails().map((mail) => mail.folderPath).filter(Boolean))
  }

  function selectOnlyMail(index: number) {
    const mail = mails.value[index]
    if (!mail?.id?.trim()) return
    selectedMailIds.value = new Set([mail.id])
    selectedMailIndex.value = index
    lastSelectedMailIndex = index
  }

  function firstSelectedMailIndex(selection: Set<string>) {
    return mails.value.findIndex((item) => item.id && selection.has(item.id))
  }

  function applyExplorerMailSelection(index: number, event?: MouseEvent) {
    const mail = mails.value[index]
    if (!mail?.id?.trim()) return 'none'
    const next = new Set(selectedMailIds.value)

    if (event?.shiftKey && lastSelectedMailIndex >= 0) {
      const start = Math.min(lastSelectedMailIndex, index)
      const end = Math.max(lastSelectedMailIndex, index)
      for (let nextIndex = start; nextIndex <= end; nextIndex += 1) {
        const id = mails.value[nextIndex]?.id
        if (id) next.add(id)
      }
      selectedMailIds.value = next
      selectedMailIndex.value = index
      return 'range'
    }

    if (event?.ctrlKey || event?.metaKey) {
      if (next.has(mail.id)) {
        next.delete(mail.id)
        const fallbackIndex = firstSelectedMailIndex(next)
        selectedMailIndex.value = fallbackIndex >= 0 ? fallbackIndex : null
      } else {
        next.add(mail.id)
        selectedMailIndex.value = index
      }
      selectedMailIds.value = next
      lastSelectedMailIndex = index
      return 'toggle'
    }

    selectOnlyMail(index)
    return 'single'
  }

  function clearSelectedMails() {
    selectedMailIds.value = new Set()
    lastSelectedMailIndex = -1
  }

  function patchMailAttachments(payload: MailAttachmentsDto) {
    if (!payload.mailId) return
    mailAttachmentsByMailId.value = {
      ...mailAttachmentsByMailId.value,
      [payload.mailId]: payload.attachments,
    }
    completeAttachmentLoad(payload.mailId)
  }

  function patchMailConversation(payload: MailConversationDto) {
    if (!payload.mailId) return
    mailConversationsByMailId.value = {
      ...mailConversationsByMailId.value,
      [payload.mailId]: payload,
    }
    completeConversationLoad(payload.mailId)
  }

  function collectExistingFolderCounts() {
    const counts = new Map<string, number>()
    function visit(items: FolderTreeNode[]) {
      for (const folder of items) {
        counts.set(folder.folderPath, folder.itemCount)
        visit(folder.subFolders)
      }
    }
    visit(folders.value)
    return counts
  }

  async function loadCachedFolders(options: { preserveExistingCounts?: boolean } = {}) {
    const snapshot = await outlookApi.getFolders()
    if (options.preserveExistingCounts) {
      const existingCounts = collectExistingFolderCounts()
      snapshot.folders = snapshot.folders.map((folder) => {
        const previousCount = existingCounts.get(folder.folderPath) ?? 0
        return previousCount > 0 && folder.itemCount === 0
          ? { ...folder, itemCount: previousCount }
          : folder
      })
    }
    folderStores.value = snapshot.stores
    folders.value = buildFolderTree(snapshot)
    selectDefaultFolder()
  }

  async function requestFolders(force = false) {
    if (outlookBusy.value && !force) return
    loadingFolders.value = true
    try {
      const response = await outlookApi.requestFolders()
      await waitForRequest(response)
      await loadCachedFolders()
      initialFoldersFetchCompleted = true
      updateOutlookFirstLoadCompleted()
      loadingFolders.value = false
      requestLoading.value = false
    } catch {
      initialFoldersFetchCompleted = true
      updateOutlookFirstLoadCompleted()
      loadingFolders.value = false
    }
  }

  function folderStore(folder: FolderTreeNode) {
    return folderStores.value.find((store) => store.storeId === folder.storeId)
  }

  function findLoadedInboxFolder() {
    const inboxes = folderOptions.value.filter((folder) => isMailSelectableFolder(folder) && folderType(folder.name) === 'inbox')
    return (
      inboxes.find((folder) => folderStore(folder)?.storeKind?.toLowerCase() === 'ost')
      ?? inboxes.find((folder) => ['exchange', 'ost'].includes(folderStore(folder)?.storeKind?.toLowerCase() ?? ''))
      ?? inboxes[0]
      ?? null
    )
  }

  function findPreferredInboxFolder() {
    return findLoadedInboxFolder() ?? folderOptions.value.find(isMailSelectableFolder) ?? null
  }

  function findPreferredInboxRootFolder() {
    return (
      visibleFolders.value.find((folder) => folderStore(folder)?.storeKind?.toLowerCase() === 'ost')
      ?? visibleFolders.value.find((folder) => ['exchange', 'ost'].includes(folderStore(folder)?.storeKind?.toLowerCase() ?? ''))
      ?? visibleFolders.value[0]
      ?? null
    )
  }

  async function ensureStartupInboxFolderLoaded() {
    if (findLoadedInboxFolder()) return
    const root = findPreferredInboxRootFolder()
    if (!root?.hasChildren || root.childrenLoaded) return

    loadingFolders.value = true
    try {
      const response = await outlookApi.requestFolderChildren({
        storeId: root.storeId,
        parentEntryId: root.entryId,
        parentFolderPath: root.folderPath,
        maxDepth: 1,
        maxChildren: 50,
      })
      await waitForRequest(response)
      await loadCachedFolders({ preserveExistingCounts: true })
    } finally {
      loadingFolders.value = false
    }
  }

  function expandFolderAncestors(folderPath: string) {
    const next = new Set(expandedFolders.value)
    let current = folderOptions.value.find((folder) => folder.folderPath === folderPath)
    while (current?.parentFolderPath) {
      next.add(current.parentFolderPath)
      current = folderOptions.value.find((folder) => folder.folderPath === current?.parentFolderPath)
    }
    expandedFolders.value = next
  }

  function selectDefaultFolder() {
    if (selectedFolderPath.value || folderOptions.value.length === 0) return
    const folder = findPreferredInboxFolder()
    selectedFolderPath.value = folder?.folderPath ?? ''
    if (selectedFolderPath.value) expandFolderAncestors(selectedFolderPath.value)
  }

  function selectInboxFolder() {
    const folder = findPreferredInboxFolder()
    selectedFolderPath.value = folder?.folderPath ?? ''
    if (selectedFolderPath.value) expandFolderAncestors(selectedFolderPath.value)
    selectedMailIndex.value = null
  }

  async function toggleFolder(path: string) {
    if (outlookBusy.value) return
    const next = new Set(expandedFolders.value)
    if (next.has(path)) next.delete(path)
    else {
      next.add(path)
      const folder = findFolderByPath(folders.value, path)
      if (folder?.hasChildren && !folder.childrenLoaded) {
        loadingFolders.value = true
        try {
          const response = await outlookApi.requestFolderChildren({
            storeId: folder.storeId,
            parentEntryId: folder.entryId,
            parentFolderPath: folder.folderPath,
            maxDepth: 1,
            maxChildren: 50,
          })
          await waitForRequest(response)
          await loadCachedFolders({ preserveExistingCounts: true })
        } finally {
          loadingFolders.value = false
        }
      }
    }
    expandedFolders.value = next
  }

  function selectFolder(path: string) {
    if (outlookBusy.value) return
    const folder = folderOptions.value.find((item) => item.folderPath === path)
    if (!folder || !isMailSelectableFolder(folder)) {
      void toggleFolder(path)
      return
    }
    selectedFolderPath.value = path
    selectedMailIndex.value = null
    scheduleMailFetch()
  }

  function scheduleMailFetch() {
    cancelScheduledMailFetch()
    if (!selectedFolderPath.value) return
    const folder = folderOptions.value.find((item) => item.folderPath === selectedFolderPath.value)
    if (!folder || !isMailSelectableFolder(folder)) return
    scheduledMailFetchAt.value = Date.now() + mailFetchDelayMs
    mailFetchCountdownTick.value = Date.now()
    mailFetchCountdownIntervalId = window.setInterval(() => {
      mailFetchCountdownTick.value = Date.now()
    }, mailFetchCountdownTickMs)
    mailFetchTimeoutId = window.setTimeout(() => {
      void requestMails()
    }, mailFetchDelayMs)
  }

  function cancelScheduledMailFetch() {
    if (mailFetchTimeoutId) window.clearTimeout(mailFetchTimeoutId)
    if (mailFetchCountdownIntervalId) window.clearInterval(mailFetchCountdownIntervalId)
    mailFetchTimeoutId = 0
    mailFetchCountdownIntervalId = 0
    scheduledMailFetchAt.value = 0
  }

  function openFolderContextMenu(payload: { path: string; x: number; y: number }) {
    if (outlookBusy.value) return
    const folder = folderOptions.value.find((item) => item.folderPath === payload.path)
    if (!folder || !isMailSelectableFolder(folder)) return
    selectedFolderPath.value = payload.path
    selectedMailIndex.value = null
    folderContextMenu.value = {
      visible: true,
      x: payload.x,
      y: payload.y,
      folderPath: payload.path,
    }
  }

  function closeFolderContextMenu() {
    folderContextMenu.value.visible = false
  }

  async function createFolderFromContext() {
    beginCreateFolder(folderContextMenu.value.folderPath)
    closeFolderContextMenu()
  }

  async function deleteFolderFromContext() {
    const targetPath = folderContextMenu.value.folderPath
    closeFolderContextMenu()
    await deleteFolder(targetPath)
  }

  async function fetchMailsFromContext() {
    const folder = folderOptions.value.find((item) => item.folderPath === folderContextMenu.value.folderPath)
    if (!folder || !isMailSelectableFolder(folder)) {
      closeFolderContextMenu()
      return
    }
    selectedFolderPath.value = folderContextMenu.value.folderPath
    closeFolderContextMenu()
    await requestMails()
  }

  async function loadCachedMails(fallbackFolderPath = '') {
    const items = await outlookApi.getMails()
    setMails(items)
    fetchedMailFolderPath.value = inferMailFolderPath(items, fallbackFolderPath)
  }

  async function loadCachedMailSearchResults() {
    setMailSearchResults(await outlookApi.getMailSearchResults())
  }

  async function loadRequestMailItems(response: { requestId?: string; request?: string }) {
    const requestId = requestIdFromResponse(response)
    const endpoint = fetchResultEndpoint(response)
    const items: MailItemDto[] = []
    let cursor = ''
    do {
      const state = await outlookApi.fetchResult<{ mails?: unknown[] }>(endpoint, {
        requestId,
        cursor,
        take: 100,
      })
      items.push(...normalizeMailItems(state.data?.mails))
      cursor = state.next.cursor
      if (!state.next.hasMore) break
    } while (cursor)
    return items
  }

  async function loadCachedRules() {
    rules.value = await outlookApi.getRules()
  }

  async function loadCachedCategories() {
    categories.value = await outlookApi.getCategories()
  }

  async function loadCachedCalendar() {
    calendarEvents.value = await outlookApi.getCalendar()
  }

  async function requestMails(force = false) {
    cancelScheduledMailFetch()
    const selectedFolder = folderOptions.value.find((folder) => folder.folderPath === selectedFolderPath.value)
    if ((outlookBusy.value && !force) || !selectedFolderPath.value || !selectedFolder || !isMailSelectableFolder(selectedFolder)) {
      if (!selectedFolderPath.value && !initialMailsFetchCompleted) {
        initialMailsFetchCompleted = true
        updateOutlookFirstLoadCompleted()
        loadingMails.value = false
      }
      return
    }
    loadingMails.value = true
    mailListMode.value = 'folder'
    pendingMailFolderPath.value = selectedFolderPath.value
    selectedMailIndex.value = null
    clearSelectedMails()
    try {
      const receivedTo = new Date()
      const receivedFrom = new Date(receivedTo.getTime() - mailLookbackHours.value * 60 * 60 * 1000)
      const response = await outlookApi.requestFolderMails({
        folderPath: selectedFolderPath.value,
        includeSubFolders: false,
        receivedFrom: receivedFrom.toISOString(),
        receivedTo: receivedTo.toISOString(),
        maxCount: mailCount.value,
      })
      await waitForRequest(response)
      const items = await loadRequestMailItems(response)
      setMails(items)
      fetchedMailFolderPath.value = inferMailFolderPath(items, pendingMailFolderPath.value)
      lastMailFetchAt.value = new Date()
      pendingMailFolderPath.value = ''
      initialMailsFetchCompleted = true
      updateOutlookFirstLoadCompleted()
      loadingMails.value = false
    } catch {
      pendingMailFolderPath.value = ''
      initialMailsFetchCompleted = true
      updateOutlookFirstLoadCompleted()
      loadingMails.value = false
    }
  }

  function localDateTimeToIso(value: string) {
    return value ? new Date(value).toISOString() : undefined
  }

  function selectedStoreIdForSearch() {
    const selectedFolder = folderOptions.value.find((folder) => folder.folderPath === selectedFolderPath.value)
    return selectedFolder?.storeId ?? ''
  }

  function searchScopeLabel(scopeMode: typeof mailSearchDraft.value.scopeMode, storeId: string, scopeFolderPaths: string[]) {
    if (scopeMode === 'global') return '全部信箱'
    if (scopeMode === 'selected_folder') return scopeFolderPaths[0] ? `${folderNameForPath(scopeFolderPaths[0])} + 子資料夾` : '目前資料夾未選擇'
    const store = folderStores.value.find((item) => item.storeId === storeId)
    return searchStoreLabel(store, storeId)
  }

  function searchDateLabel(value: string) {
    return value ? value.replace('T', ' ') : ''
  }

  function searchReceivedCondition(from: string, to: string) {
    if (from && to) return `${searchDateLabel(from)} <= 時間 <= ${searchDateLabel(to)}`
    if (from) return `時間 >= ${searchDateLabel(from)}`
    if (to) return `時間 <= ${searchDateLabel(to)}`
    return ''
  }

  function buildMailSearchSummary(storeId: string, scopeFolderPaths: string[]) {
    const draft = mailSearchDraft.value
    const keyword = draft.keyword.trim()
    const receivedCondition = searchReceivedCondition(draft.receivedFrom, draft.receivedTo)
    const fieldLabels: Record<string, string> = {
      subject: '標題',
      sender: '寄件者',
      body: '內容',
    }
    const textFields = draft.textFields.map(field => fieldLabels[field] ?? field).join('、') || '標題'
    const summary = [
      { label: '範圍', value: searchScopeLabel(draft.scopeMode, storeId, scopeFolderPaths), tone: 'info' },
      { label: '文字範圍', value: textFields, tone: 'info' },
    ]
    if (draft.categoryNames.length > 0) summary.push({ label: '分類', value: draft.categoryNames.join('、'), tone: 'info' })
    if (draft.hasAttachments !== undefined) summary.push({ label: '附件', value: draft.hasAttachments ? '包含附件' : '不含附件', tone: 'info' })
    if (draft.flagState !== 'any') summary.push({ label: '旗標', value: draft.flagState === 'flagged' ? '有旗標' : '無旗標', tone: 'info' })
    if (draft.readState !== 'any') summary.push({ label: '狀態', value: draft.readState === 'unread' ? '未讀' : '已讀', tone: 'info' })
    if (receivedCondition) summary.unshift({ label: '時間', value: receivedCondition, tone: 'active' })
    if (keyword) summary.unshift({ label: '文字', value: `包含 "${keyword}"`, tone: 'active' })
    else summary.push({ label: '文字', value: '未使用', tone: 'muted' })
    return summary as typeof activeMailSearchSummary.value
  }

  async function requestMailSearch() {
    if (loadingMailSearch.value) return
    const searchId = window.crypto?.randomUUID?.() ?? `${Date.now()}`
    const scopeFolderPaths = mailSearchDraft.value.scopeMode === 'selected_folder' && selectedFolderPath.value
      ? [selectedFolderPath.value]
      : []
    const storeId = mailSearchDraft.value.scopeMode === 'global' ? '' : selectedStoreIdForSearch()
    activeMailSearchSummary.value = buildMailSearchSummary(storeId, scopeFolderPaths)
    loadingMailSearch.value = true
    mailSearchProgress.value = null
    mailListMode.value = 'search'
    mailSearchResults.value = []
    collapsedSearchResultStores.value = new Set()
    collapsedSearchResultFolders.value = new Set()
    selectedMailIndex.value = null
    clearSelectedMails()
    try {
      const response = await outlookApi.requestMailSearch({
        searchId,
        storeId,
        scopeFolderPaths,
        allowGlobalScope: mailSearchDraft.value.scopeMode === 'global',
        includeSubFolders: true,
        keyword: mailSearchDraft.value.keyword,
        textFields: mailSearchDraft.value.textFields,
        categoryNames: mailSearchDraft.value.categoryNames,
        hasAttachments: mailSearchDraft.value.hasAttachments,
        flagState: mailSearchDraft.value.flagState,
        readState: mailSearchDraft.value.readState,
        receivedFrom: localDateTimeToIso(mailSearchDraft.value.receivedFrom),
        receivedTo: localDateTimeToIso(mailSearchDraft.value.receivedTo),
      })
      await waitForRequest(response)
      await loadCachedFolders({ preserveExistingCounts: true })
      try {
        const result = await outlookApi.fetchResult<{ searchId?: string }>(fetchResultEndpoint(response), {
          requestId: requestIdFromResponse(response),
          take: 1,
        })
        mailSearchProgress.value = result.data?.searchId
          ? await outlookApi.getMailSearchProgress(result.data.searchId)
          : null
      } catch {
        // Search progress 不是每個失敗路徑都一定會留下 snapshot。
      }
      setMailSearchResults(await loadRequestMailItems(response))
      loadingMailSearch.value = false
    } catch {
      loadingMailSearch.value = false
    }
  }

  function showFolderMails() {
    mailListMode.value = 'folder'
    selectedMailIndex.value = null
    lastSelectedMailIndex = -1
    if (mailListNeedsFetch.value && !outlookBusy.value) scheduleMailFetch()
  }

  function openMailDialog(index: number) {
    const mail = mails.value[index]
    if (!mail) return
    selectOnlyMail(index)
    mailDialogIndex.value = index
    mailDialogMailId.value = mail.id
    mailDialogHtml.value = false
    mailDialogVisible.value = true
    void requestMailBody(mail)
    void requestMailAttachments(mail)
    void requestMailConversation(mail)
  }

  function closeMailDialog() {
    mailDialogVisible.value = false
    mailDialogMailId.value = ''
    mailDialogHtml.value = false
  }

  async function requestRules() {
    if (loadingRules.value) return
    if (outlookBusy.value) return
    loadingRules.value = true
    try {
      const response = await outlookApi.requestRules()
      await waitForRequest(response)
      await loadCachedRules()
      loadingRules.value = false
    } catch {
      loadingRules.value = false
    }
  }

  function splitRuleInput(value: string) {
    return value
      .split(/[\n,;]+/)
      .map((item) => item.trim())
      .filter(Boolean)
  }

  function parseRuleSummaryValue(summary: string, key: string) {
    const marker = `${key}=`
    const index = summary.indexOf(marker)
    if (index < 0) return ''
    const rest = summary.slice(index + marker.length)
    const nextPart = rest.indexOf('; ')
    return (nextPart >= 0 ? rest.slice(0, nextPart) : rest).trim()
  }

  function parseRuleSummaryList(summary: string, key: string) {
    return parseRuleSummaryValue(summary, key)
      .split(',')
      .map((item) => item.trim())
      .filter(Boolean)
  }

  function firstRuleSummaryValue(summaries: string[], prefix: string, key: string) {
    const summary = summaries.find((item) => item.toLowerCase().startsWith(prefix.toLowerCase()))
    return summary ? parseRuleSummaryValue(summary, key) : ''
  }

  function firstRuleSummaryList(summaries: string[], prefix: string, key: string) {
    const summary = summaries.find((item) => item.toLowerCase().startsWith(prefix.toLowerCase()))
    return summary ? parseRuleSummaryList(summary, key) : []
  }

  function resetRuleDraft(rule: OutlookRuleDto | null = null) {
    if (!rule) {
      ruleDraft.value = {
        ruleName: '',
        storeId: '',
        originalRuleName: '',
        originalExecutionOrder: undefined,
        ruleType: 'receive',
        enabled: true,
        subjectContains: '',
        bodyContains: '',
        senderAddressContains: '',
        categories: [],
        hasAttachment: 'any',
        moveToFolderPath: '',
        assignCategories: [],
        markAsTask: false,
        stopProcessingMoreRules: true,
      }
      selectedRuleIndex.value = null
      return
    }

    ruleDraft.value = {
      ruleName: rule.name,
      storeId: rule.storeId,
      originalRuleName: rule.name,
      originalExecutionOrder: rule.executionOrder,
      ruleType: rule.ruleType?.toLowerCase() === 'send' ? 'send' : 'receive',
      enabled: rule.enabled,
      subjectContains: firstRuleSummaryList(rule.conditions, 'Subject:', 'Text').join(', '),
      bodyContains: firstRuleSummaryList(rule.conditions, 'Body:', 'Text').join(', '),
      senderAddressContains: firstRuleSummaryList(rule.conditions, 'SenderAddress:', 'Address').join(', '),
      categories: firstRuleSummaryList(rule.conditions, 'Category:', 'Categories'),
      hasAttachment: rule.conditions.some((condition) => condition.toLowerCase().startsWith('hasattachment:')) ? 'yes' : 'any',
      moveToFolderPath: firstRuleSummaryValue(rule.actions, 'MoveToFolder:', 'FolderPath'),
      assignCategories: firstRuleSummaryList(rule.actions, 'AssignToCategory:', 'Categories'),
      markAsTask: rule.actions.some((action) => action.toLowerCase().includes('task')),
      stopProcessingMoreRules: rule.actions.some((action) => action.toLowerCase().includes('stop')),
    }
  }

  function editRule(index: number) {
    const rule = rules.value[index]
    if (!rule) return
    selectedRuleIndex.value = index
    resetRuleDraft(rule)
  }

  function buildRulePayload(operation: 'upsert' | 'delete' | 'set_enabled', source: RuleDraft = ruleDraft.value): OutlookRuleCommandRequest {
    const draft = source
    const hasAttachment = draft.hasAttachment === 'any' ? undefined : draft.hasAttachment === 'yes'
    return {
      operation,
      storeId: draft.storeId,
      ruleName: draft.ruleName.trim() || draft.originalRuleName.trim(),
      originalRuleName: draft.originalRuleName.trim(),
      originalExecutionOrder: draft.originalExecutionOrder,
      ruleType: draft.ruleType,
      enabled: draft.enabled,
      executionOrder: draft.originalExecutionOrder,
      conditions: {
        subjectContains: splitRuleInput(draft.subjectContains),
        bodyContains: splitRuleInput(draft.bodyContains),
        senderAddressContains: splitRuleInput(draft.senderAddressContains),
        categories: draft.categories,
        hasAttachment,
      },
      actions: {
        moveToFolderPath: draft.moveToFolderPath,
        assignCategories: draft.assignCategories,
        markAsTask: draft.markAsTask,
        stopProcessingMoreRules: draft.stopProcessingMoreRules,
      },
    }
  }

  function buildRuleOperationDraft(rule: OutlookRuleDto, enabled = rule.enabled): RuleDraft {
    return {
      ruleName: rule.name,
      storeId: rule.storeId,
      originalRuleName: rule.name,
      originalExecutionOrder: rule.executionOrder,
      ruleType: rule.ruleType?.toLowerCase() === 'send' ? 'send' : 'receive',
      enabled,
      subjectContains: '',
      bodyContains: '',
      senderAddressContains: '',
      categories: [],
      hasAttachment: 'any',
      moveToFolderPath: '',
      assignCategories: [],
      markAsTask: false,
      stopProcessingMoreRules: false,
    }
  }

  async function saveRule() {
    if (outlookBusy.value || !ruleDraft.value.ruleName.trim()) return false
    const hasCondition = splitRuleInput(ruleDraft.value.subjectContains).length > 0
      || splitRuleInput(ruleDraft.value.bodyContains).length > 0
      || splitRuleInput(ruleDraft.value.senderAddressContains).length > 0
      || ruleDraft.value.categories.length > 0
      || ruleDraft.value.hasAttachment !== 'any'
    const hasAction = Boolean(ruleDraft.value.moveToFolderPath)
      || ruleDraft.value.assignCategories.length > 0
      || ruleDraft.value.markAsTask
      || ruleDraft.value.stopProcessingMoreRules
    if (!hasCondition || !hasAction) {
      ElMessage.warning('Rule 需要至少一個條件與一個動作。')
      return false
    }

    return await runMailOperation(
      () => outlookApi.requestManageRule(buildRulePayload('upsert')),
      async () => {
        await loadCachedRules()
        resetRuleDraft()
      },
    )
  }

  async function deleteRule(rule = selectedRule.value) {
    if (!rule || outlookBusy.value) return
    const confirmed = window.confirm(`刪除 Outlook rule「${rule.name}」？`)
    if (!confirmed) return
    const payload = buildRulePayload('delete', buildRuleOperationDraft(rule))
    await runMailOperation(
      () => outlookApi.requestManageRule(payload),
      async () => {
        await loadCachedRules()
        resetRuleDraft()
      },
    )
  }

  async function toggleRuleEnabled(rule: OutlookRuleDto, enabled: boolean) {
    if (outlookBusy.value) return
    const payload = buildRulePayload('set_enabled', buildRuleOperationDraft(rule, enabled))
    await runMailOperation(
      () => outlookApi.requestManageRule(payload),
      async () => {
        await loadCachedRules()
      },
    )
  }

  async function requestCategories(force = false) {
    if (outlookBusy.value && !force) return
    loadingCategories.value = true
    try {
      const response = await outlookApi.requestCategories()
      await waitForRequest(response)
      await loadCachedCategories()
      initialCategoriesFetchCompleted = true
      updateOutlookFirstLoadCompleted()
      loadingCategories.value = false
      requestLoading.value = false
    } catch {
      initialCategoriesFetchCompleted = true
      updateOutlookFirstLoadCompleted()
      loadingCategories.value = false
    }
  }

  async function requestSignalRPing() {
    if (loadingSignalRPing.value) return
    loadingSignalRPing.value = true
    try {
      const response = await outlookApi.requestSignalRPing()
      await waitForRequest(response)
    } finally {
      loadingSignalRPing.value = false
    }
  }

  async function runStartupOutlookSync() {
    if (startupSyncStarted) return
    if (unmounted) return
    startupSyncStarted = true
    await requestFolders(true)
    if (!initialFoldersFetchCompleted) {
      initialFoldersFetchCompleted = true
      updateOutlookFirstLoadCompleted()
      loadingFolders.value = false
    }
    if (unmounted) return
    await ensureStartupInboxFolderLoaded()
    if (unmounted) return
    selectInboxFolder()

    await requestCategories(true)

    if (unmounted) return
    if (!selectedFolderPath.value) selectInboxFolder()
    mailLookbackHours.value = 168
    mailCount.value = 30
    await requestMails(true)
  }

  async function requestCalendar() {
    if (outlookBusy.value) return
    loadingCalendar.value = true
    try {
      const start = monthStart(calendarMonthDate.value)
      const end = monthEndExclusive(calendarMonthDate.value)
      const response = await outlookApi.requestCalendar({
        daysForward: Math.ceil((end.getTime() - start.getTime()) / 86400000),
        startDate: toDateKey(start),
        endDate: toDateKey(end),
      })
      await waitForRequest(response)
      await loadCachedCalendar()
      loadingCalendar.value = false
    } catch {
      loadingCalendar.value = false
    }
  }

  async function changeCalendarMonth(offset: number) {
    if (outlookBusy.value) return
    calendarMonthDate.value = addMonths(calendarMonthDate.value, offset)
    selectedCalendarEvent.value = null
    await requestCalendar()
  }

  async function goToCurrentCalendarMonth() {
    if (outlookBusy.value) return
    calendarMonthDate.value = monthStart(new Date())
    selectedCalendarEvent.value = null
    await requestCalendar()
  }

  function selectCalendarEvent(event: CalendarEventDto) {
    selectedCalendarEvent.value = event
  }

  async function runMailOperation(action: () => Promise<unknown>, afterSuccess?: () => Promise<void>) {
    if (outlookBusy.value) return false
    window.clearTimeout(requestTimeoutId)
    activeRequestId = ''
    requestLoading.value = true
    try {
      const response = await action()
      if (isRequestResponse(response)) {
        activeRequestId = requestIdFromResponse(response)
        if (!['accepted'].includes(response.state)) {
          requestLoading.value = false
          activeRequestId = ''
          return false
        }
        await waitForRequest(response)
      }
      if (afterSuccess) await afterSuccess()
      completeRequest(activeRequestId)
      return true
    } catch {
      requestLoading.value = false
      activeRequestId = ''
      window.clearTimeout(requestTimeoutId)
      return false
    }
  }

  function isRequestResponse(value: unknown): value is { requestId?: string; request?: string; state: string; data?: unknown } {
    const response = value as { requestId?: unknown; request?: unknown; state?: unknown; data?: unknown }
    return typeof response?.requestId === 'string'
      && typeof response?.request === 'string'
      && typeof response?.state === 'string'
      && response.data !== undefined
  }

  function requestIdFromResponse(response: { requestId?: string }) {
    return response.requestId || ''
  }

  async function waitForRequest(response: { requestId?: string; request?: string }, timeoutMs = 120000) {
    const requestId = requestIdFromResponse(response)
    if (!requestId) return
    const endpoint = fetchResultEndpoint(response)
    const started = Date.now()
    while (!unmounted && Date.now() - started < timeoutMs) {
      try {
        const state = await outlookApi.fetchResult(endpoint, {
          requestId,
          take: 1,
        })
        if (state.state === 'completed') return
        if (state.state && !['accepted', 'running'].includes(state.state)) {
          throw new Error(state.message || 'Outlook operation failed')
        }
      } catch (error) {
        if (error instanceof Error && error.message !== 'Request failed: 404') throw error
      }
      await new Promise((resolve) => window.setTimeout(resolve, 300))
    }
    throw new Error('Outlook operation timed out')
  }

  function fetchResultEndpoint(response: { request?: string }) {
    const request = response.request || ''
    return request.startsWith('request-')
      ? request.replace('request-', 'fetch-result-')
      : 'fetch-result-mails'
  }

  function completeRequest(requestId = '') {
    if (requestId && activeRequestId && requestId !== activeRequestId) return
    requestLoading.value = false
    activeRequestId = ''
    window.clearTimeout(requestTimeoutId)
  }

  function completeMailBodyLoad(mailId: string) {
    const next = new Set(loadingMailBodyIds.value)
    next.delete(mailId)
    loadingMailBodyIds.value = next
    const timeoutId = mailBodyTimeoutIds.get(mailId)
    if (timeoutId) window.clearTimeout(timeoutId)
    mailBodyTimeoutIds.delete(mailId)
  }

  function clearMailBodyLoads() {
    loadingMailBodyIds.value = new Set()
    for (const timeoutId of mailBodyTimeoutIds.values()) window.clearTimeout(timeoutId)
    mailBodyTimeoutIds.clear()
  }

  function attachmentKey(mailId: string, attachmentId: string) {
    return `${mailId}\n${attachmentId}`
  }

  function completeAttachmentLoad(mailId: string) {
    const next = new Set(loadingAttachmentMailIds.value)
    next.delete(mailId)
    loadingAttachmentMailIds.value = next
    const timeoutId = attachmentTimeoutIds.get(mailId)
    if (timeoutId) window.clearTimeout(timeoutId)
    attachmentTimeoutIds.delete(mailId)
  }

  function completeAttachmentExport(mailId: string, attachmentId: string) {
    const key = attachmentKey(mailId, attachmentId)
    const next = new Set(exportingAttachmentIds.value)
    next.delete(key)
    exportingAttachmentIds.value = next
    const timeoutId = attachmentTimeoutIds.get(key)
    if (timeoutId) window.clearTimeout(timeoutId)
    attachmentTimeoutIds.delete(key)
  }

  function clearAttachmentLoads() {
    loadingAttachmentMailIds.value = new Set()
    exportingAttachmentIds.value = new Set()
    for (const timeoutId of attachmentTimeoutIds.values()) window.clearTimeout(timeoutId)
    attachmentTimeoutIds.clear()
  }

  function completeConversationLoad(mailId: string) {
    const next = new Set(loadingConversationMailIds.value)
    next.delete(mailId)
    loadingConversationMailIds.value = next
  }

  function clearConversationLoads() {
    loadingConversationMailIds.value = new Set()
  }

  function isMailBodyLoading(mail: MailItemDto) {
    return Boolean(mail.id && loadingMailBodyIds.value.has(mail.id))
  }

  function mailHasBody(mail: MailItemDto) {
    return Boolean(mail.body || mail.bodyHtml)
  }

  function isAttachmentListLoading(mail: MailItemDto) {
    return Boolean(mail.id && loadingAttachmentMailIds.value.has(mail.id))
  }

  function isAttachmentExporting(mail: MailItemDto, attachment: MailAttachmentDto) {
    return Boolean(mail.id && exportingAttachmentIds.value.has(attachmentKey(mail.id, attachment.attachmentId)))
  }

  function isConversationLoading(mail: MailItemDto) {
    return Boolean(mail.id && loadingConversationMailIds.value.has(mail.id))
  }

  async function requestMailBody(mail: MailItemDto) {
    if (!mail.id?.trim() || mailHasBody(mail) || isMailBodyLoading(mail)) return
    loadingMailBodyIds.value = new Set(loadingMailBodyIds.value).add(mail.id)
    try {
      const response = await outlookApi.requestMailBody({
        mailId: mail.id,
        folderPath: mail.folderPath,
      })
      await waitForRequest(response)
      await refreshMailSnapshotsForDetail(mail.id)
      completeMailBodyLoad(mail.id)
    } catch {
      completeMailBodyLoad(mail.id)
    }
  }

  async function refreshMailSnapshotsForDetail(mailId: string) {
    const mode = mailListMode.value
    await Promise.allSettled([loadCachedMails(), loadCachedMailSearchResults()])
    mailListMode.value = mode
    const list = mode === 'search' ? mailSearchResults.value : folderMails.value
    const nextIndex = list.findIndex((item) => item.id === mailId)
    if (nextIndex >= 0) selectedMailIndex.value = nextIndex
  }

  async function requestMailAttachments(mail: MailItemDto) {
    if (!mail.id?.trim() || isAttachmentListLoading(mail) || mailAttachmentsByMailId.value[mail.id]) return
    loadingAttachmentMailIds.value = new Set(loadingAttachmentMailIds.value).add(mail.id)
    try {
      const response = await outlookApi.requestMailAttachments({
        mailId: mail.id,
        folderPath: mail.folderPath,
      })
      await waitForRequest(response)
      patchMailAttachments(await outlookApi.getMailAttachments(mail.id))
    } catch {
      completeAttachmentLoad(mail.id)
    }
  }

  async function requestMailConversation(mail: MailItemDto) {
    if (!mail.id?.trim() || isConversationLoading(mail) || mailConversationsByMailId.value[mail.id]) return
    loadingConversationMailIds.value = new Set(loadingConversationMailIds.value).add(mail.id)
    try {
      const response = await outlookApi.requestMailConversation({
        mailId: mail.id,
        folderPath: mail.folderPath,
        maxCount: 100,
        includeBody: true,
      })
      await waitForRequest(response)
      const requestId = requestIdFromResponse(response)
      const items: MailItemDto[] = []
      let cursor = ''
      let conversation: MailConversationDto | null = null
      do {
        const state = await outlookApi.fetchResult<{
          mailId?: string
          folderPath?: string
          conversationId?: string
          conversationTopic?: string
          mails?: unknown[]
        }>(fetchResultEndpoint(response), {
          requestId,
          cursor,
          take: 100,
        })
        const data = state.data ?? {}
        conversation = {
          mailId: data.mailId || mail.id,
          folderPath: data.folderPath || mail.folderPath,
          conversationId: data.conversationId || mail.conversationId,
          conversationTopic: data.conversationTopic || mail.conversationTopic || mail.subject,
          mails: items,
        }
        items.push(...normalizeMailItems(data.mails))
        cursor = state.next.cursor
        if (!state.next.hasMore) break
      } while (cursor)

      if (conversation) {
        conversation.mails = items
        patchMailConversation(conversation)
      } else {
        patchMailConversation(await outlookApi.getMailConversation(mail.id))
      }
    } catch {
      completeConversationLoad(mail.id)
    }
  }

  async function exportMailAttachment(mail: MailItemDto, attachment: MailAttachmentDto) {
    if (!mail.id?.trim() || !attachment.attachmentId || isAttachmentExporting(mail, attachment)) return
    const key = attachmentKey(mail.id, attachment.attachmentId)
    const exportAttachmentId = attachment.index > 0 ? String(attachment.index) : attachment.attachmentId
    exportingAttachmentIds.value = new Set(exportingAttachmentIds.value).add(key)
    try {
      const response = await outlookApi.requestExportMailAttachment({
        mailId: mail.id,
        folderPath: mail.folderPath,
        attachmentId: exportAttachmentId,
        index: attachment.index,
        name: attachment.name,
        fileName: attachment.fileName,
        displayName: attachment.displayName,
      })
      await waitForRequest(response)
      patchMailAttachments(await outlookApi.getMailAttachments(mail.id))
      completeAttachmentExport(mail.id, attachment.attachmentId)
    } catch {
      completeAttachmentExport(mail.id, attachment.attachmentId)
    }
  }

  async function openExportedAttachment(attachment: MailAttachmentDto) {
    if (!attachment.exportedAttachmentId) return
    await outlookApi.openExportedAttachment({ exportedAttachmentId: attachment.exportedAttachmentId })
  }

  async function applyMailProperties(mail: MailItemDto) {
    if (!mail.id?.trim()) return
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

  async function upsertCategory(name: string, color: string, shortcutKey = '') {
    const categoryName = name.trim()
    if (!categoryName || outlookBusy.value) return
    requestLoading.value = true
    try {
      const response = await outlookApi.requestUpsertCategory({
        name: categoryName,
        color: color || 'olCategoryColorNone',
        colorValue: categoryColorValue(color || 'olCategoryColorNone'),
        shortcutKey,
      })
      await waitForRequest(response)
      await loadCachedCategories()
      requestLoading.value = false
    } catch {
      requestLoading.value = false
    }
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

  function beginCreateFolder(parentPath: string) {
    if (!parentPath || outlookBusy.value) return
    creatingFolderParentPath.value = parentPath
    creatingFolderName.value = ''
    const next = new Set(expandedFolders.value)
    next.add(parentPath)
    expandedFolders.value = next
  }

  function cancelCreateFolder() {
    creatingFolderParentPath.value = ''
    creatingFolderName.value = ''
  }

  function deletedFolderForPath(path: string) {
    if (!path) return null
    const exactFolder = folderOptions.value.find((folder) => folder.folderPath === path)
    const storeId = exactFolder?.storeId
    return folderOptions.value.find((folder) =>
      folder.folderType === 'Deleted'
      && (!storeId || folder.storeId === storeId)
      && (path === folder.folderPath || path.startsWith(`${folder.folderPath}/`))
    ) ?? null
  }

  function isInDeletedFolder(path: string) {
    return deletedFolderForPath(path) !== null
  }

  async function createFolder(parentPath = creatingFolderParentPath.value, name = creatingFolderName.value) {
    const folderName = name.trim()
    if (!parentPath || !folderName || outlookBusy.value) return
    const parent = folderOptions.value.find((folder) => folder.folderPath === parentPath)
    if (!parent || !isMailSelectableFolder(parent)) return
    requestLoading.value = true
    try {
      const response = await outlookApi.requestCreateFolder({
        parentFolderPath: parentPath,
        name: folderName,
      })
      await waitForRequest(response)
      await loadCachedFolders()
      cancelCreateFolder()
      requestLoading.value = false
    } catch {
      requestLoading.value = false
    }
  }

  async function deleteFolder(targetPath: string) {
    if (!targetPath || outlookBusy.value) return
    const targetFolder = folderOptions.value.find((folder) => folder.folderPath === targetPath)
    if (!targetFolder || !isMailSelectableFolder(targetFolder)) return
    const targetName = folderOptions.value.find((folder) => folder.folderPath === targetPath)?.label.trim() ?? targetPath
    if (isInDeletedFolder(targetPath)) {
      ElMessage.warning(manualOutlookDeleteMessage)
      return
    }
    const confirmed = window.confirm(`將 Folder「${targetName}」移到 Outlook 刪除資料夾？`)
    if (!confirmed) return
    requestLoading.value = true
    try {
      const response = await outlookApi.requestDeleteFolder({
        folderPath: targetPath,
      })
      await waitForRequest(response)
      await loadCachedFolders()
      if (selectedFolderPath.value === targetPath) {
        selectedFolderPath.value = folderOptions.value[0]?.folderPath ?? ''
      }
      requestLoading.value = false
    } catch {
      requestLoading.value = false
    }
  }

  function selectMail(index: number, event?: MouseEvent) {
    const selectionMode = applyExplorerMailSelection(index, event)
    if (selectionMode === 'none') return
  }

  function captureMailListSnapshot() {
    return {
      folderMails: [...folderMails.value],
      mailSearchResults: [...mailSearchResults.value],
      selectedMailIds: new Set(selectedMailIds.value),
      selectedMailIndex: selectedMailIndex.value,
      lastSelectedMailIndex,
    }
  }

  function restoreMailListSnapshot(snapshot: ReturnType<typeof captureMailListSnapshot>) {
    folderMails.value = snapshot.folderMails
    mailSearchResults.value = snapshot.mailSearchResults
    selectedMailIds.value = snapshot.selectedMailIds
    selectedMailIndex.value = snapshot.selectedMailIndex
    lastSelectedMailIndex = snapshot.lastSelectedMailIndex
  }

  function hideMovedMails(mailIds: string[]) {
    const movedIds = new Set(mailIds.filter(Boolean))
    if (movedIds.size === 0) return
    if (mailListMode.value === 'search') {
      mailSearchResults.value = mailSearchResults.value.filter((mail) => !movedIds.has(mail.id))
    } else {
      folderMails.value = folderMails.value.filter((mail) => !movedIds.has(mail.id))
    }

    selectedMailIds.value = new Set([...selectedMailIds.value].filter((id) => !movedIds.has(id)))
    const nextIndex = firstSelectedMailIndex(selectedMailIds.value)
    selectedMailIndex.value = nextIndex >= 0 ? nextIndex : null
    if (selectedMailIndex.value === null) lastSelectedMailIndex = -1
  }

  function restoreFailedMailMove(snapshot: ReturnType<typeof captureMailListSnapshot>) {
    restoreMailListSnapshot(snapshot)
    ElMessage.error('移動郵件失敗，已還原畫面。')
  }

  async function moveMailToFolder(mail: MailItemDto, destinationFolderPath: string) {
    if (!mail.id?.trim() || !destinationFolderPath || destinationFolderPath === mail.folderPath) return
    const snapshot = captureMailListSnapshot()
    hideMovedMails([mail.id])
    const succeeded = await runMailOperation(
      () =>
        outlookApi.requestMoveMail({
          mailId: mail.id,
          sourceFolderPath: mail.folderPath,
          destinationFolderPath,
        }),
      async () => {
        await loadCachedFolders()
      },
    )
    if (!succeeded) restoreFailedMailMove(snapshot)
  }

  async function moveSelectedMailsToFolder(destinationFolderPath: string) {
    const selected = selectedBulkMoveMails()
    if (selected.length === 0 || !destinationFolderPath || outlookBusy.value) return
    const sourceFolderPaths = [...new Set(selected.map((mail) => mail.folderPath).filter(Boolean))]
    const snapshot = captureMailListSnapshot()
    hideMovedMails(selected.map((mail) => mail.id))
    const succeeded = await runMailOperation(
      () =>
        outlookApi.requestMoveMails({
          mailIds: selected.map((mail) => mail.id),
          sourceFolderPath: sourceFolderPaths.length === 1 ? sourceFolderPaths[0] : '',
          sourceFolderPaths,
          destinationFolderPath,
          continueOnError: true,
        }),
      async () => {
        await loadCachedFolders()
      },
    )
    if (succeeded) clearSelectedMails()
    else restoreFailedMailMove(snapshot)
  }

  function currentDragMailCount(mailId: string) {
    return selectedMailIds.value.has(mailId) && selectedMailIds.value.size > 1
      ? selectedMailIds.value.size
      : 1
  }

  function setMailDragPreview(event: DragEvent, count: number) {
    if (!event.dataTransfer) return
    const preview = document.createElement('div')
    preview.className = 'mail-drag-preview'
    preview.textContent = `移動 ${count} 封郵件`
    document.body.appendChild(preview)
    event.dataTransfer.setDragImage(preview, 18, 18)
    window.setTimeout(() => preview.remove(), 0)
  }

  async function deleteMail(mail: MailItemDto) {
    if (!mail?.id?.trim() || outlookBusy.value) return
    if (isInDeletedFolder(mail.folderPath)) {
      ElMessage.warning(manualOutlookDeleteMessage)
      return
    }
    const deletedFolder = deletedFolderForPath(mail.folderPath) ?? folderOptions.value.find((folder) => folder.folderType === 'Deleted')
    const targetName = deletedFolder?.label.trim() || '刪除的郵件 / Deleted Items'
    const confirmed = window.confirm(`將郵件「${mail.subject || mail.id}」移到「${targetName}」？`)
    if (!confirmed) return
    await runMailOperation(
      () =>
        outlookApi.requestDeleteMail({
          mailId: mail.id,
          folderPath: mail.folderPath,
        }),
      async () => {
        await Promise.allSettled([loadCachedMails(), loadCachedMailSearchResults(), loadCachedFolders()])
      },
    )
  }

  function startMailDrag(mail: MailItemDto, index: number, event: DragEvent) {
    if (!mail.id?.trim()) {
      event.preventDefault()
      return
    }
    if (outlookBusy.value) return
    if (!selectedMailIds.value.has(mail.id)) selectOnlyMail(index)
    draggedMailId.value = mail.id
    event.dataTransfer?.setData('text/plain', mail.id)
    if (event.dataTransfer) {
      const count = currentDragMailCount(mail.id)
      event.dataTransfer.effectAllowed = 'move'
      event.dataTransfer.setData('text/x-smartoffice-mail-count', String(count))
      setMailDragPreview(event, count)
    }
  }

  function clearMailDrag() {
    draggedMailId.value = ''
    dragOverFolderPath.value = ''
  }

  function setDragOverFolder(path: string) {
    if (!draggedMailId.value || outlookBusy.value) return
    const folder = folderOptions.value.find((item) => item.folderPath === path)
    if (!folder || !isMailSelectableFolder(folder)) return
    dragOverFolderPath.value = path
  }

  async function moveDraggedMail(destinationFolderPath: string) {
    const mailId = draggedMailId.value
    clearMailDrag()
    if (!mailId || outlookBusy.value) return
    const destination = folderOptions.value.find((item) => item.folderPath === destinationFolderPath)
    if (!destination || !isMailSelectableFolder(destination)) return
    const mail = mails.value.find((item) => item.id === mailId)
    if (!mail) return
    if (selectedMailIds.value.has(mailId) && selectedMailIds.value.size > 1) {
      await moveSelectedMailsToFolder(destinationFolderPath)
      return
    }
    await moveMailToFolder(mail, destinationFolderPath)
  }

  async function loadChat() {
    chatMessages.value = await outlookApi.getChat()
    await scrollChatToBottom()
  }

  async function sendChat() {
    const text = chatText.value.trim()
    if (!text) return
    chatText.value = ''
    await outlookApi.sendChat({ source: 'web', text })
    await loadChat()
  }

  async function refreshAdminData() {
    const [status, logs] = await Promise.all([
      outlookApi.getAdminStatus(),
      outlookApi.getAdminLogs(),
    ])
    addinStatus.value = status
    addinLogs.value = logs
  }

  async function loadAttachmentExportSettings() {
    const settings = await outlookApi.getAttachmentExportSettings()
    attachmentExportSettings.value = settings
    attachmentExportRootDraft.value = settings.rootPath
  }

  async function saveAttachmentExportSettings() {
    if (savingAttachmentExportSettings.value) return
    savingAttachmentExportSettings.value = true
    try {
      const settings = await outlookApi.updateAttachmentExportSettings({
        rootPath: attachmentExportRootDraft.value,
      })
      attachmentExportSettings.value = settings
      attachmentExportRootDraft.value = settings.rootPath
    } finally {
      savingAttachmentExportSettings.value = false
    }
  }

  async function resetAttachmentExportRoot() {
    attachmentExportRootDraft.value = attachmentExportSettings.value.defaultRootPath
    await saveAttachmentExportSettings()
  }

  async function switchView(view: AppView) {
    if (outlookDependentViewsLocked.value && ['search', 'rules', 'chat', 'calendar'].includes(view)) return
    activeView.value = view
    if (view === 'outlook') {
      mailListMode.value = 'folder'
      selectedMailIndex.value = null
      lastSelectedMailIndex = -1
      if (mailListNeedsFetch.value && !outlookBusy.value) scheduleMailFetch()
      return
    }
    if (view === 'search') {
      cancelScheduledMailFetch()
      mailListMode.value = 'search'
      selectedMailIndex.value = null
      lastSelectedMailIndex = -1
    }
    if (view === 'rules' && rules.value.length === 0) void requestRules()
  }

  async function scrollChatToBottom() {
    await nextTick()
    if (chatPanelRef.value) chatPanelRef.value.scrollTop = chatPanelRef.value.scrollHeight
  }

  async function connectSignalR() {
    connection = new signalR.HubConnectionBuilder()
      .withUrl('/hub/notifications')
      .withAutomaticReconnect()
      .build()

    connection.onreconnecting(() => {
      signalRState.value = 'reconnecting'
    })
    connection.onreconnected(() => {
      signalRState.value = 'connected'
      void refreshAdminData()
    })
    connection.onclose(() => {
      signalRState.value = 'disconnected'
    })
    connection.on('AddinStatus', (status: AddinStatusDto) => {
      addinStatus.value = status
    })
    connection.on('AddinLog', (logs: AddinLogEntry[]) => {
      addinLogs.value = logs
    })

    try {
      await connection.start()
      signalRState.value = 'connected'
    } catch {
      signalRState.value = 'disconnected'
    }
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

  onMounted(async () => {
    unmounted = false
    window.addEventListener('click', closeFolderContextMenu)
    await connectSignalR()
    await Promise.allSettled([
      loadCachedFolders(),
      loadCachedMails(),
      loadCachedMailSearchResults(),
      loadCachedRules(),
      loadCachedCategories(),
      loadCachedCalendar(),
      loadChat(),
      refreshAdminData(),
      loadAttachmentExportSettings(),
    ])
    void runStartupOutlookSync()
  })

  onBeforeUnmount(() => {
    unmounted = true
    window.removeEventListener('click', closeFolderContextMenu)
    window.clearTimeout(requestTimeoutId)
    cancelScheduledMailFetch()
    clearMailBodyLoads()
    clearAttachmentLoads()
    clearConversationLoads()
    void connection?.stop()
  })

  return {
    activeView,
    activeMailPropertySections,
    addCategoryToMasterList,
    addinLogs,
    addinStatus,
    attachmentExportRootDraft,
    attachmentExportSettings,
    addMailCategoryDraft,
    applyMailProperties,
    calendarEvents,
    calendarMonthLabel,
    calendarWeekdays,
    calendarWeeks,
    cancelCreateFolder,
    categories,
    categoryManagerVisible,
    categoryColorOptions,
    categoryColorStyle,
    categoryTagStyle,
    categoryCreateColor,
    categoryCreateDraft,
    changeCalendarMonth,
    chatMessages,
    chatPanelRef,
    chatText,
    clearMailDrag,
    clearSelectedMails,
    contextFolderName,
    createFolder,
    createFolderFromContext,
    creatingFolderName,
    creatingFolderParentPath,
    deleteFolderFromContext,
    deleteMail,
    dragOverFolderPath,
    draggedMailId,
    expandedFolders,
    exportMailAttachment,
    fetchMailsFromContext,
    flagDisplayLabel,
    flagIntervalOptions,
    flagTagType,
    folderContextMenu,
    folderStores,
    loadingCalendar,
    loadingCategories,
    loadAttachmentExportSettings,
    loadingFolders,
    loadingMails,
    loadingMailSearch,
    loadingSignalRPing,
    mailCount,
    mailHtmlSandbox,
    mailListMode,
    mailSearchDraft,
    mailSearchProgress,
    mailSearchProgressText,
    mailSearchSummaryItems,
    searchResultGroups,
    searchResultRows,
    searchResultViewMode,
    mailSearchResults,
    isAttachmentExporting,
    isAttachmentListLoading,
    isConversationLoading,
    isMailBodyLoading,
    mailHasBody,
    dialogMail,
    dialogMailAttachments,
    dialogMailConversation,
    dialogMailConversationItems,
    dialogMailFolderName,
    dialogMailHasIdentity,
    dialogLoading,
    mailDialogHtml,
    mailDialogVisible,
    mailPropertiesDraft,
    mailPropertiesChanged,
    mailLookbackHours,
    mailStats,
    masterCategoryListExpanded,
    mails,
    moveDraggedMail,
    navOptions,
    openFolderContextMenu,
    openExportedAttachment,
    openMailDialog,
    closeMailDialog,
    operationLoading: requestLoading,
    outlookBusy,
    outlookBusyText,
    openCategoryManager,
    refreshAdminData,
    requestCalendar,
    requestCategories,
    requestFolders,
    requestRules,
    requestSignalRPing,
    requestMails,
    requestMailSearch,
    resetMailPropertiesDraft,
    resetRuleDraft,
    resetAttachmentExportRoot,
    removeMailCategoryDraft,
    saveAttachmentExportSettings,
    savingAttachmentExportSettings,
    fetchedMailFolderName,
    mailListNeedsFetch,
    mailFetchCountdownText,
    showMailFetchWarning,
    mailFetchStatusText,
    selectedFolderName,
    selectedCalendarEvent,
    selectedFolderPath,
    selectedMail,
    selectedMailCategories,
    selectedMailIndex,
    selectedMailIds,
    selectedRule,
    selectedRuleIndex,
    selectFolder,
    selectCalendarEvent,
    selectMail,
    sendChat,
    saveRule,
    deleteRule,
    editRule,
    showFolderMails,
    goToCurrentCalendarMonth,
    setDragOverFolder,
    setMailFlagDraft,
    signalRState,
    splitCategories,
    startMailDrag,
    switchView,
    toggleFolder,
    toggleSearchResultFolder,
    toggleSearchResultStore,
    updateCategoryColor,
    toggleRuleEnabled,
    visibleFolders,
    folderOptions,
    ruleDraft,
    ruleDraftIsEditing,
    rules,
    loadingRules,
    visibleMasterCategories,
    hiddenMasterCategoryCount,
    toggleMasterCategoryList,
    flagEditorVisible,
  }
}

export type OutlookDashboardState = ReturnType<typeof useOutlookDashboard>
