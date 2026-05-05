import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import * as signalR from '@microsoft/signalr'
import {
  normalizeExportedMailAttachment,
  normalizeMailAttachments,
  normalizeMailBody,
  normalizeMailItem,
  normalizeMailSearchSliceResult,
  normalizeMailSearchProgress,
  normalizeMailItems,
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
  FolderSyncBatchDto,
  FolderSyncBeginDto,
  FolderSyncCompleteDto,
  FolderTreeNode,
  ExportedMailAttachmentDto,
  MailAttachmentDto,
  MailAttachmentsDto,
  MailBodyDto,
  MailItemDto,
  MailPropertiesDraft,
  MailSearchCompleteDto,
  MailSearchProgressDto,
  MailPropertiesCommandRequest,
  OutlookCommandResult,
  OutlookStoreDto,
  OutlookCategoryDto,
  OutlookRuleDto,
  SignalRState,
} from '../models/outlook'
import {
  categoryColorOptions,
  categoryColorStyle,
  categoryColorValue,
  categoryOptionColor,
  categoryTextColor,
} from '../utils/categoryColors'
import { applyFolderBatch, buildFolderTree, collectFolderOptions, findFolderByPath, folderType, visibleRootFolders } from '../utils/folders'
import {
  addMonths,
  dateInputToIso,
  defaultFlagRequest,
  flagDisplayLabel,
  flagIntervalOptions,
  flagTagType,
  isDefaultFlagRequest,
  mergeStores,
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

export function useOutlookDashboard() {
  const activeView = ref<AppView>('outlook')
  const signalRState = ref<SignalRState>('disconnected')
  const folders = ref<FolderTreeNode[]>([])
  const folderStores = ref<OutlookStoreDto[]>([])
  const folderMails = ref<MailItemDto[]>([])
  const mailSearchResults = ref<MailItemDto[]>([])
  const mailListMode = ref<'folder' | 'search'>('folder')
  const rules = ref<OutlookRuleDto[]>([])
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
  const selectedMailHtml = ref(false)
  const activeMailPropertySections = ref(['set-mail-properties'])
  const expandedFolders = ref<Set<string>>(new Set())
  const openMailIndexes = ref<Set<number>>(new Set())
  const htmlMailIndexes = ref<Set<number>>(new Set())
  const loadingMailBodyIds = ref<Set<string>>(new Set())
  const mailAttachmentsByMailId = ref<Record<string, MailAttachmentDto[]>>({})
  const loadingAttachmentMailIds = ref<Set<string>>(new Set())
  const exportingAttachmentIds = ref<Set<string>>(new Set())
  const mailRange = ref('1m')
  const mailCount = ref(30)
  const loadingMailSearch = ref(false)
  const activeMailSearchId = ref('')
  const mailSearchProgress = ref<MailSearchProgressDto | null>(null)
  const mailSearchDraft = ref({
    keyword: '',
    matchMode: 'contains' as 'contains' | 'exact' | 'fuzzy',
    fields: ['subject'] as string[],
    receivedFrom: '',
    receivedTo: '',
    scopeMode: 'selected_folder' as 'selected_folder' | 'selected_store' | 'global',
    maxCount: 50,
  })
  const chatText = ref('')
  const loadingFolders = ref(false)
  const loadingMails = ref(true)
  const loadingRules = ref(false)
  const loadingCategories = ref(true)
  const loadingCalendar = ref(false)
  const loadingSignalRPing = ref(false)
  const operationLoading = ref(false)
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
  let activeOperationCommandId = ''
  let operationTimeoutId = 0
  let activeMailSearchCommandId = ''
  let mailSearchTimeoutId = 0
  const mailBodyCommandIds = new Map<string, string>()
  const mailBodyTimeoutIds = new Map<string, number>()
  const attachmentCommandIds = new Map<string, { type: 'fetch' | 'export'; mailId: string; attachmentId?: string }>()
  const attachmentTimeoutIds = new Map<string, number>()

  const visibleFolders = computed(() => visibleRootFolders(folders.value))
  const mails = computed(() => mailListMode.value === 'search' ? mailSearchResults.value : folderMails.value)

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
    return loadingFolders.value || loadingMails.value || loadingRules.value || loadingCategories.value || loadingCalendar.value || operationLoading.value
  })

  const outlookBusyText = computed(() => {
    if (loadingFolders.value) return 'Outlook folder 同步中...'
    if (loadingMails.value) return 'Outlook 郵件抓取中...'
    if (loadingRules.value) return 'Outlook rule 同步中...'
    if (loadingCategories.value) return 'Outlook category 同步中...'
    if (loadingCalendar.value) return 'Outlook calendar 同步中...'
    if (operationLoading.value) return 'Outlook 操作執行中...'
    return ''
  })

  const folderOptions = computed(() => collectFolderOptions(visibleFolders.value))

  const calendarWeekdays = ['日', '一', '二', '三', '四', '五', '六']

  const calendarMonthLabel = computed(() => {
    return calendarMonthDate.value.toLocaleDateString('zh-TW', { year: 'numeric', month: 'long' })
  })

  const calendarWeeks = computed(() => {
    const first = monthStart(calendarMonthDate.value)
    const gridStart = new Date(first)
    gridStart.setDate(first.getDate() - first.getDay())
    const todayKey = toDateKey(new Date())

    return Array.from({ length: 6 }, (_, weekIndex) => {
      const weekStart = new Date(gridStart)
      weekStart.setDate(gridStart.getDate() + weekIndex * 7)
      const weekEnd = new Date(weekStart)
      weekEnd.setDate(weekStart.getDate() + 6)
      const days = Array.from({ length: 7 }, (_, dayIndex) => {
        const date = new Date(gridStart)
        date.setDate(gridStart.getDate() + weekIndex * 7 + dayIndex)
        const key = toDateKey(date)
        return {
          key,
          date,
          dayNumber: date.getDate(),
          inMonth: date.getMonth() === calendarMonthDate.value.getMonth(),
          isToday: key === todayKey,
        }
      })

      const segments = calendarEvents.value
        .map((event) => calendarEventSegment(event, weekStart, weekEnd))
        .filter((segment): segment is NonNullable<typeof segment> => Boolean(segment))
        .sort((a, b) => new Date(a.event.start).getTime() - new Date(b.event.start).getTime())

      return {
        key: days.map((day) => day.key).join('-'),
        days,
        segments,
      }
    })
  })

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

  const selectedMailFolderName = computed(() => {
    return selectedMail.value?.folderPath ? folderNameForPath(selectedMail.value.folderPath) : '未選擇'
  })

  const mailListNeedsFetch = computed(() => {
    if (mailListMode.value === 'search') return false
    return Boolean(selectedFolderPath.value && selectedFolderPath.value !== fetchedMailFolderPath.value)
  })

  const contextFolderName = computed(() => {
    return folderOptions.value.find((folder) => folder.folderPath === folderContextMenu.value.folderPath)?.label.trim() ?? '未選擇'
  })

  const selectedMail = computed(() => {
    if (selectedMailIndex.value === null) return null
    return mails.value[selectedMailIndex.value] ?? null
  })

  const selectedMailIsOpen = computed(() => {
    return selectedMailIndex.value !== null && openMailIndexes.value.has(selectedMailIndex.value)
  })

  const selectedMailCategories = computed(() => {
    if (!selectedMail.value?.categories) return []
    return selectedMail.value.categories
      .split(',')
      .map((category) => category.trim())
      .filter(Boolean)
  })

  const selectedMailHasIdentity = computed(() => Boolean(selectedMail.value?.id?.trim()))

  const selectedMailAttachments = computed(() => {
    return selectedMail.value?.id ? mailAttachmentsByMailId.value[selectedMail.value.id] ?? [] : []
  })

  const mailPropertiesChanged = computed(() => {
    if (!selectedMail.value) return false
    return JSON.stringify(buildMailPropertiesPayload(selectedMail.value, buildMailPropertiesDraft(selectedMail.value)))
      !== JSON.stringify(buildMailPropertiesPayload(selectedMail.value, mailPropertiesDraft.value))
  })

  function folderNameForPath(path: string) {
    if (!path) return '未選擇'
    return folderOptions.value.find((folder) => folder.folderPath === path)?.label.trim() ?? path
  }

  function calendarEventSegment(event: CalendarEventDto, weekStart: Date, weekEnd: Date) {
    const start = new Date(event.start)
    const end = new Date(event.end)
    if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return null
    const eventEnd = new Date(end)
    if (eventEnd.getTime() > start.getTime()) eventEnd.setMilliseconds(eventEnd.getMilliseconds() - 1)
    const startDay = new Date(start.getFullYear(), start.getMonth(), start.getDate())
    const endDay = new Date(eventEnd.getFullYear(), eventEnd.getMonth(), eventEnd.getDate())
    if (endDay < weekStart || startDay > weekEnd) return null

    const segmentStart = startDay < weekStart ? weekStart : startDay
    const segmentEnd = endDay > weekEnd ? weekEnd : endDay
    const startColumn = Math.floor((segmentStart.getTime() - weekStart.getTime()) / 86400000) + 1
    const span = Math.floor((segmentEnd.getTime() - segmentStart.getTime()) / 86400000) + 1

    return {
      event,
      startColumn,
      span,
      isStart: startDay >= weekStart,
      isEnd: endDay <= weekEnd,
      isMultiDay: endDay.getTime() > startDay.getTime(),
    }
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

  function setMails(items: MailItemDto[], preferredMailId = selectedMail.value?.id ?? '') {
    const wasOpen = selectedMailIndex.value !== null && openMailIndexes.value.has(selectedMailIndex.value)
    clearMailBodyLoads()
    clearAttachmentLoads()
    mailAttachmentsByMailId.value = {}
    folderMails.value = items
    mailListMode.value = 'folder'

    if (items.length === 0) {
      selectedMailIndex.value = null
      openMailIndexes.value = new Set()
      htmlMailIndexes.value = new Set()
      return
    }

    const nextIndex = preferredMailId ? items.findIndex((mail) => mail.id === preferredMailId) : -1
    selectedMailIndex.value = nextIndex >= 0 ? nextIndex : 0

    if (wasOpen && selectedMailIndex.value !== null) {
      openMailIndexes.value = new Set([selectedMailIndex.value])
    }
  }

  function patchMail(nextMail: MailItemDto) {
    if (!nextMail.id) return
    patchMailInList(folderMails, nextMail)
    patchMailInList(mailSearchResults, nextMail)
  }

  function patchMailBody(body: MailBodyDto) {
    if (!body.mailId) return
    patchMailBodyInList(folderMails, body)
    patchMailBodyInList(mailSearchResults, body)
    completeMailBodyLoad(body.mailId)
  }

  function patchMailInList(target: typeof folderMails, nextMail: MailItemDto) {
    const index = target.value.findIndex((mail) => mail.id === nextMail.id)
    if (index < 0) return
    const items = [...target.value]
    items[index] = {
      ...nextMail,
      body: nextMail.body || items[index].body,
      bodyHtml: nextMail.bodyHtml || items[index].bodyHtml,
    }
    target.value = items
  }

  function patchMailBodyInList(target: typeof folderMails, body: MailBodyDto) {
    const index = target.value.findIndex((mail) => mail.id === body.mailId)
    if (index < 0) return
    const items = [...target.value]
    items[index] = {
      ...items[index],
      body: body.body,
      bodyHtml: body.bodyHtml,
      folderPath: body.folderPath || items[index].folderPath,
    }
    target.value = items
  }

  function patchMailAttachments(payload: MailAttachmentsDto) {
    if (!payload.mailId) return
    mailAttachmentsByMailId.value = {
      ...mailAttachmentsByMailId.value,
      [payload.mailId]: payload.attachments,
    }
    completeAttachmentLoad(payload.mailId)
  }

  function patchExportedAttachment(payload: ExportedMailAttachmentDto) {
    if (!payload.mailId || !payload.attachmentId) return
    const attachments = mailAttachmentsByMailId.value[payload.mailId] ?? []
    const matchedAttachment = attachments.find((attachment) => isSameAttachment(attachment, payload))
    mailAttachmentsByMailId.value = {
      ...mailAttachmentsByMailId.value,
      [payload.mailId]: attachments.map((attachment) =>
        isSameAttachment(attachment, payload)
          ? {
              ...attachment,
              isExported: true,
              exportedAttachmentId: payload.exportedAttachmentId,
              exportedPath: payload.exportedPath,
              size: payload.size || attachment.size,
            }
          : attachment,
      ),
    }
    if (matchedAttachment) completeAttachmentExport(payload.mailId, matchedAttachment.attachmentId)
    completeAttachmentExport(payload.mailId, payload.attachmentId)
  }

  function isSameAttachment(attachment: MailAttachmentDto, exported: ExportedMailAttachmentDto) {
    return (
      attachment.attachmentId === exported.attachmentId
      || Boolean(attachment.id && attachment.id === exported.id)
      || Boolean(attachment.index > 0 && attachment.index === exported.index)
      || Boolean(attachment.name && attachment.name === exported.name)
    )
  }

  async function loadCachedFolders() {
    const snapshot = await outlookApi.getFolders()
    folderStores.value = snapshot.stores
    folders.value = buildFolderTree(snapshot)
    selectDefaultFolder()
  }

  async function requestFolders(force = false) {
    if (outlookBusy.value && !force) return
    loadingFolders.value = true
    try {
      await outlookApi.requestFolders()
      await loadCachedFolders()
      initialFoldersFetchCompleted = true
      loadingFolders.value = false
      operationLoading.value = false
    } catch {
      initialFoldersFetchCompleted = true
      loadingFolders.value = false
    }
  }

  function folderStore(folder: FolderTreeNode) {
    return folderStores.value.find((store) => store.storeId === folder.storeId)
  }

  function findLoadedInboxFolder() {
    const inboxes = folderOptions.value.filter((folder) => folderType(folder.name) === 'inbox')
    return (
      inboxes.find((folder) => folderStore(folder)?.storeKind?.toLowerCase() === 'ost')
      ?? inboxes.find((folder) => ['exchange', 'ost'].includes(folderStore(folder)?.storeKind?.toLowerCase() ?? ''))
      ?? inboxes[0]
      ?? null
    )
  }

  function findPreferredInboxFolder() {
    return findLoadedInboxFolder() ?? folderOptions.value[0] ?? null
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
      await outlookApi.requestFolderChildren({
        storeId: root.storeId,
        parentEntryId: root.entryId,
        parentFolderPath: root.folderPath,
        maxDepth: 1,
        maxChildren: 50,
      })
      await loadCachedFolders()
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
    selectedMailHtml.value = false
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
          await outlookApi.requestFolderChildren({
            storeId: folder.storeId,
            parentEntryId: folder.entryId,
            parentFolderPath: folder.folderPath,
            maxDepth: 1,
            maxChildren: 50,
          })
        } finally {
          loadingFolders.value = false
        }
      }
    }
    expandedFolders.value = next
  }

  function selectFolder(path: string) {
    if (outlookBusy.value) return
    selectedFolderPath.value = path
    selectedMailIndex.value = null
    selectedMailHtml.value = false
  }

  function openFolderContextMenu(payload: { path: string; x: number; y: number }) {
    if (outlookBusy.value) return
    selectFolder(payload.path)
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
    selectedFolderPath.value = folderContextMenu.value.folderPath
    closeFolderContextMenu()
    await requestMails()
  }

  async function loadCachedMails() {
    const items = await outlookApi.getMails()
    setMails(items)
    fetchedMailFolderPath.value = inferMailFolderPath(items)
  }

  async function loadCachedMailSearchResults() {
    mailSearchResults.value = await outlookApi.getMailSearchResults()
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
    if ((outlookBusy.value && !force) || !selectedFolderPath.value) {
      if (!selectedFolderPath.value && !initialMailsFetchCompleted) {
        initialMailsFetchCompleted = true
        loadingMails.value = false
      }
      return
    }
    loadingMails.value = true
    mailListMode.value = 'folder'
    pendingMailFolderPath.value = selectedFolderPath.value
    openMailIndexes.value = new Set()
    htmlMailIndexes.value = new Set()
    selectedMailIndex.value = null
    selectedMailHtml.value = false
    try {
      await outlookApi.requestMails({
        folderPath: selectedFolderPath.value,
        range: mailRange.value,
        maxCount: mailCount.value,
      })
      await loadCachedMails()
      pendingMailFolderPath.value = ''
      initialMailsFetchCompleted = true
      loadingMails.value = false
    } catch {
      pendingMailFolderPath.value = ''
      initialMailsFetchCompleted = true
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

  async function requestMailSearch() {
    if (loadingMailSearch.value) return
    const searchId = window.crypto?.randomUUID?.() ?? `${Date.now()}`
    const scopeFolderPaths = mailSearchDraft.value.scopeMode === 'selected_folder' && selectedFolderPath.value
      ? [selectedFolderPath.value]
      : []
    const storeId = mailSearchDraft.value.scopeMode === 'global' ? '' : selectedStoreIdForSearch()
    loadingMailSearch.value = true
    activeMailSearchId.value = searchId
    mailSearchProgress.value = null
    mailListMode.value = 'search'
    mailSearchResults.value = []
    selectedMailIndex.value = null
    openMailIndexes.value = new Set()
    htmlMailIndexes.value = new Set()
    selectedMailHtml.value = false
    try {
      const response = await outlookApi.requestMailSearch({
        searchId,
        storeId,
        scopeFolderPaths,
        includeSubFolders: true,
        keyword: mailSearchDraft.value.keyword,
        matchMode: mailSearchDraft.value.matchMode,
        fields: mailSearchDraft.value.fields,
        receivedFrom: localDateTimeToIso(mailSearchDraft.value.receivedFrom),
        receivedTo: localDateTimeToIso(mailSearchDraft.value.receivedTo),
        maxCount: mailSearchDraft.value.maxCount,
      })
      activeMailSearchCommandId = response.commandId
      try {
        mailSearchProgress.value = await outlookApi.getMailSearchProgressByCommandId(response.commandId)
      } catch {
        // SignalR progress may arrive before this HTTP lookup is available.
      }
      window.clearTimeout(mailSearchTimeoutId)
      mailSearchTimeoutId = window.setTimeout(() => {
        loadingMailSearch.value = false
        activeMailSearchCommandId = ''
      }, 120000)
    } catch {
      loadingMailSearch.value = false
    }
  }

  function showFolderMails() {
    mailListMode.value = 'folder'
    selectedMailIndex.value = null
    openMailIndexes.value = new Set()
    htmlMailIndexes.value = new Set()
  }

  function applyMailSearchSliceResult(item: unknown) {
    const batch = normalizeMailSearchSliceResult(item)
    if (activeMailSearchId.value && batch.searchId && batch.searchId !== activeMailSearchId.value) return
    if (batch.searchId) activeMailSearchId.value = batch.searchId
    const current = batch.reset ? [] : mailSearchResults.value
    const byId = new Map(current.map((mail) => [mail.id, mail]))
    for (const mail of batch.mails) byId.set(mail.id, mail)
    mailSearchResults.value = [...byId.values()]
    mailListMode.value = 'search'
    if (selectedMailIndex.value === null && mailSearchResults.value.length > 0) selectedMailIndex.value = 0
    if (batch.isFinal) loadingMailSearch.value = false
  }

  async function requestRules() {
    if (outlookBusy.value) return
    loadingRules.value = true
    try {
      await outlookApi.requestRules()
      await loadCachedRules()
      loadingRules.value = false
    } catch {
      loadingRules.value = false
    }
  }

  async function requestCategories(force = false) {
    if (outlookBusy.value && !force) return
    loadingCategories.value = true
    try {
      await outlookApi.requestCategories()
      await loadCachedCategories()
      initialCategoriesFetchCompleted = true
      loadingCategories.value = false
      operationLoading.value = false
    } catch {
      initialCategoriesFetchCompleted = true
      loadingCategories.value = false
    }
  }

  async function requestSignalRPing() {
    if (loadingSignalRPing.value) return
    loadingSignalRPing.value = true
    try {
      await outlookApi.requestSignalRPing()
    } finally {
      loadingSignalRPing.value = false
    }
  }

  async function waitForNotificationSignalRConnected(timeoutMs = 12000) {
    const started = Date.now()
    while (!unmounted && signalRState.value !== 'connected' && Date.now() - started < timeoutMs) {
      await new Promise((resolve) => window.setTimeout(resolve, 100))
    }
    return signalRState.value === 'connected'
  }

  async function runStartupOutlookSync() {
    if (startupSyncStarted) return
    const connected = await waitForNotificationSignalRConnected()
    if (!connected || unmounted) return
    startupSyncStarted = true
    await requestFolders(true)
    if (!initialFoldersFetchCompleted) {
      initialFoldersFetchCompleted = true
      loadingFolders.value = false
    }
    if (unmounted) return
    await ensureStartupInboxFolderLoaded()
    if (unmounted) return
    selectInboxFolder()

    await requestCategories(true)

    if (unmounted) return
    if (!selectedFolderPath.value) selectInboxFolder()
    mailRange.value = '1m'
    mailCount.value = 30
    await requestMails(true)
  }

  async function requestCalendar() {
    if (outlookBusy.value) return
    loadingCalendar.value = true
    try {
      const start = monthStart(calendarMonthDate.value)
      const end = monthEndExclusive(calendarMonthDate.value)
      await outlookApi.requestCalendar({
        daysForward: Math.ceil((end.getTime() - start.getTime()) / 86400000),
        startDate: toDateKey(start),
        endDate: toDateKey(end),
      })
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

  async function runMailOperation(action: () => Promise<unknown>) {
    if (outlookBusy.value) return
    window.clearTimeout(operationTimeoutId)
    activeOperationCommandId = ''
    operationLoading.value = true
    try {
      const response = await action()
      if (!operationLoading.value) return
      if (isCommandDispatchResponse(response)) {
        activeOperationCommandId = response.commandId
        if (response.status === 'completed' || response.status === 'mocked') {
          completeOperation(response.commandId)
          return
        }
        if (response.status !== 'dispatched' && response.status !== 'mocked') {
          operationLoading.value = false
          activeOperationCommandId = ''
          return
        }
      }
      operationTimeoutId = window.setTimeout(() => {
        operationLoading.value = false
        activeOperationCommandId = ''
      }, 30000)
    } catch {
      operationLoading.value = false
      activeOperationCommandId = ''
      window.clearTimeout(operationTimeoutId)
    }
  }

  function isCommandDispatchResponse(value: unknown): value is { commandId: string; status: string } {
    const response = value as { commandId?: unknown; status?: unknown }
    return typeof response?.commandId === 'string' && typeof response?.status === 'string'
  }

  function completeOperation(commandId = '') {
    if (commandId && activeOperationCommandId && commandId !== activeOperationCommandId) return
    operationLoading.value = false
    activeOperationCommandId = ''
    window.clearTimeout(operationTimeoutId)
  }

  function completeMailBodyLoad(mailId: string) {
    const next = new Set(loadingMailBodyIds.value)
    next.delete(mailId)
    loadingMailBodyIds.value = next
    const timeoutId = mailBodyTimeoutIds.get(mailId)
    if (timeoutId) window.clearTimeout(timeoutId)
    mailBodyTimeoutIds.delete(mailId)
    for (const [commandId, trackedMailId] of mailBodyCommandIds) {
      if (trackedMailId === mailId) mailBodyCommandIds.delete(commandId)
    }
  }

  function clearMailBodyLoads() {
    loadingMailBodyIds.value = new Set()
    for (const timeoutId of mailBodyTimeoutIds.values()) window.clearTimeout(timeoutId)
    mailBodyTimeoutIds.clear()
    mailBodyCommandIds.clear()
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
    for (const [commandId, tracked] of attachmentCommandIds) {
      if (tracked.type === 'fetch' && tracked.mailId === mailId) attachmentCommandIds.delete(commandId)
    }
  }

  function completeAttachmentExport(mailId: string, attachmentId: string) {
    const key = attachmentKey(mailId, attachmentId)
    const next = new Set(exportingAttachmentIds.value)
    next.delete(key)
    exportingAttachmentIds.value = next
    const timeoutId = attachmentTimeoutIds.get(key)
    if (timeoutId) window.clearTimeout(timeoutId)
    attachmentTimeoutIds.delete(key)
    for (const [commandId, tracked] of attachmentCommandIds) {
      if (tracked.type === 'export' && tracked.mailId === mailId && tracked.attachmentId === attachmentId) {
        attachmentCommandIds.delete(commandId)
      }
    }
  }

  function clearAttachmentLoads() {
    loadingAttachmentMailIds.value = new Set()
    exportingAttachmentIds.value = new Set()
    for (const timeoutId of attachmentTimeoutIds.values()) window.clearTimeout(timeoutId)
    attachmentTimeoutIds.clear()
    attachmentCommandIds.clear()
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

  async function requestMailBody(mail: MailItemDto) {
    if (!mail.id?.trim() || mailHasBody(mail) || isMailBodyLoading(mail)) return
    loadingMailBodyIds.value = new Set(loadingMailBodyIds.value).add(mail.id)
    try {
      const response = await outlookApi.requestMailBody({
        mailId: mail.id,
        folderPath: mail.folderPath,
      })
      if (response.status === 'completed' || response.status === 'mocked') {
        await loadCachedMails()
        completeMailBodyLoad(mail.id)
        return
      }
      mailBodyCommandIds.set(response.commandId, mail.id)
      const timeoutId = window.setTimeout(() => completeMailBodyLoad(mail.id), 30000)
      mailBodyTimeoutIds.set(mail.id, timeoutId)
    } catch {
      completeMailBodyLoad(mail.id)
    }
  }

  async function requestMailAttachments(mail: MailItemDto) {
    if (!mail.id?.trim() || isAttachmentListLoading(mail) || mailAttachmentsByMailId.value[mail.id]) return
    loadingAttachmentMailIds.value = new Set(loadingAttachmentMailIds.value).add(mail.id)
    try {
      const response = await outlookApi.requestMailAttachments({
        mailId: mail.id,
        folderPath: mail.folderPath,
      })
      if (mailAttachmentsByMailId.value[mail.id]) {
        completeAttachmentLoad(mail.id)
        return
      }
      attachmentCommandIds.set(response.commandId, { type: 'fetch', mailId: mail.id })
      const timeoutId = window.setTimeout(() => completeAttachmentLoad(mail.id), 30000)
      attachmentTimeoutIds.set(mail.id, timeoutId)
    } catch {
      completeAttachmentLoad(mail.id)
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
      const current = mailAttachmentsByMailId.value[mail.id]?.find((item) => item.attachmentId === attachment.attachmentId)
      if (current?.isExported) {
        completeAttachmentExport(mail.id, attachment.attachmentId)
        return
      }
      attachmentCommandIds.set(response.commandId, { type: 'export', mailId: mail.id, attachmentId: attachment.attachmentId })
      const timeoutId = window.setTimeout(() => completeAttachmentExport(mail.id, attachment.attachmentId), 30000)
      attachmentTimeoutIds.set(key, timeoutId)
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
    await runMailOperation(() => outlookApi.requestUpdateMailProperties(body))
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
    operationLoading.value = true
    try {
      await outlookApi.requestUpsertCategory({
        name: categoryName,
        color: color || 'olCategoryColorNone',
        colorValue: categoryColorValue(color || 'olCategoryColorNone'),
        shortcutKey,
      })
    } catch {
      operationLoading.value = false
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

  async function createFolder(parentPath = creatingFolderParentPath.value, name = creatingFolderName.value) {
    const folderName = name.trim()
    if (!parentPath || !folderName || outlookBusy.value) return
    operationLoading.value = true
    try {
      await outlookApi.requestCreateFolder({
        parentFolderPath: parentPath,
        name: folderName,
      })
      cancelCreateFolder()
    } catch {
      operationLoading.value = false
    }
  }

  async function deleteFolder(targetPath: string) {
    if (!targetPath || outlookBusy.value) return
    const targetName = folderOptions.value.find((folder) => folder.folderPath === targetPath)?.label.trim() ?? targetPath
    const confirmed = window.confirm(`刪除 Folder「${targetName}」？`)
    if (!confirmed) return
    operationLoading.value = true
    try {
      await outlookApi.requestDeleteFolder({
        folderPath: targetPath,
      })
      if (selectedFolderPath.value === targetPath) {
        selectedFolderPath.value = folderOptions.value[0]?.folderPath ?? ''
      }
    } catch {
      operationLoading.value = false
    }
  }

  function selectMail(index: number) {
    if (selectedMailIndex.value === index) {
      const next = new Set(openMailIndexes.value)
      if (next.has(index)) next.delete(index)
      else {
        next.add(index)
        void requestMailBody(mails.value[index])
        void requestMailAttachments(mails.value[index])
      }
      openMailIndexes.value = next
      return
    }

    selectedMailIndex.value = index
    selectedMailHtml.value = false
    const next = new Set<number>()
    next.add(index)
    openMailIndexes.value = next
    void requestMailBody(mails.value[index])
    void requestMailAttachments(mails.value[index])
  }

  async function moveMailToFolder(mail: MailItemDto, destinationFolderPath: string) {
    if (!mail.id?.trim() || !destinationFolderPath || destinationFolderPath === mail.folderPath) return
    await runMailOperation(() =>
      outlookApi.requestMoveMail({
        mailId: mail.id,
        sourceFolderPath: mail.folderPath,
        destinationFolderPath,
      }),
    )
  }

  async function deleteMail(mail: MailItemDto) {
    if (!mail?.id?.trim() || outlookBusy.value) return
    const deletedFolder = folderOptions.value.find((folder) => folderType(folder.name) === 'deleted')
    const targetName = deletedFolder?.label.trim() || '刪除的郵件 / Deleted Items'
    const confirmed = window.confirm(`將郵件「${mail.subject || mail.id}」移到「${targetName}」？`)
    if (!confirmed) return
    await runMailOperation(() =>
      outlookApi.requestDeleteMail({
        mailId: mail.id,
        folderPath: mail.folderPath,
      }),
    )
  }

  function startMailDrag(mail: MailItemDto, index: number, event: DragEvent) {
    if (!mail.id?.trim()) {
      event.preventDefault()
      return
    }
    if (outlookBusy.value) return
    selectMail(index)
    draggedMailId.value = mail.id
    event.dataTransfer?.setData('text/plain', mail.id)
    if (event.dataTransfer) event.dataTransfer.effectAllowed = 'move'
  }

  function clearMailDrag() {
    draggedMailId.value = ''
    dragOverFolderPath.value = ''
  }

  function setDragOverFolder(path: string) {
    if (!draggedMailId.value || outlookBusy.value) return
    dragOverFolderPath.value = path
  }

  async function moveDraggedMail(destinationFolderPath: string) {
    const mailId = draggedMailId.value
    clearMailDrag()
    if (!mailId || outlookBusy.value) return
    const mail = mails.value.find((item) => item.id === mailId)
    if (!mail) return
    await moveMailToFolder(mail, destinationFolderPath)
  }

  function toggleMail(index: number) {
    const next = new Set(openMailIndexes.value)
    if (next.has(index)) next.delete(index)
    else next.add(index)
    openMailIndexes.value = next
  }

  function toggleMailFormat(index: number) {
    const next = new Set(htmlMailIndexes.value)
    if (next.has(index)) next.delete(index)
    else next.add(index)
    htmlMailIndexes.value = next
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
    activeView.value = view
    if (view === 'search') {
      mailListMode.value = 'search'
      selectedMailIndex.value = null
      openMailIndexes.value = new Set()
      htmlMailIndexes.value = new Set()
      selectedMailHtml.value = false
    }
    if (view === 'admin') {
      await Promise.allSettled([
        refreshAdminData(),
        loadAttachmentExportSettings(),
      ])
    }
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
      void runStartupOutlookSync()
    })
    connection.onclose(() => {
      signalRState.value = 'disconnected'
    })
    connection.on('FolderSyncStarted', (_info: FolderSyncBeginDto) => {
      loadingFolders.value = true
    })
    connection.on('FoldersPatched', (batch: FolderSyncBatchDto) => {
      if (batch.reset && batch.isFinal && batch.stores.length === 0 && batch.folders.length === 0 && folderOptions.value.length > 0) {
        loadingFolders.value = false
        completeOperation()
        return
      }

      folderStores.value = mergeStores(batch.reset ? [] : folderStores.value, batch.stores)
      folders.value = applyFolderBatch(folders.value, folderStores.value, batch)
      selectDefaultFolder()
      loadingFolders.value = !batch.isFinal
      if (batch.isFinal) completeOperation()
    })
    connection.on('FolderSyncCompleted', (_info: FolderSyncCompleteDto) => {
      initialFoldersFetchCompleted = true
      loadingFolders.value = false
      completeOperation()
    })
    connection.on('MailsUpdated', (items: unknown) => {
      const nextMails = normalizeMailItems(items)
      setMails(nextMails)
      fetchedMailFolderPath.value = pendingMailFolderPath.value || inferMailFolderPath(nextMails, fetchedMailFolderPath.value)
      pendingMailFolderPath.value = ''
      initialMailsFetchCompleted = true
      loadingMails.value = false
      completeOperation()
    })
    connection.on('MailSearchStarted', (item: unknown) => {
      const batch = normalizeMailSearchSliceResult(item)
      if (batch.searchId) activeMailSearchId.value = batch.searchId
      if (batch.reset) mailSearchResults.value = []
      mailListMode.value = 'search'
      loadingMailSearch.value = true
    })
    connection.on('MailSearchPatched', (item: unknown) => {
      applyMailSearchSliceResult(item)
    })
    connection.on('MailSearchProgress', (item: unknown) => {
      const progress = normalizeMailSearchProgress(item)
      if (activeMailSearchId.value && progress.searchId && progress.searchId !== activeMailSearchId.value) return
      mailSearchProgress.value = progress
      if (progress.status === 'running') loadingMailSearch.value = true
      if (['completed', 'failed', 'addin_unavailable'].includes(progress.status)) loadingMailSearch.value = false
    })
    connection.on('MailSearchCompleted', (info: MailSearchCompleteDto) => {
      if (activeMailSearchId.value && info.searchId && info.searchId !== activeMailSearchId.value) return
      loadingMailSearch.value = false
      activeMailSearchCommandId = ''
      window.clearTimeout(mailSearchTimeoutId)
    })
    connection.on('MailUpdated', (item: unknown) => {
      const mail = normalizeMailItem(item)
      patchMail(mail)
      if (mailHasBody(mail)) completeMailBodyLoad(mail.id)
      completeOperation()
    })
    connection.on('MailBodyUpdated', (item: unknown) => {
      patchMailBody(normalizeMailBody(item))
    })
    connection.on('MailAttachmentsUpdated', (item: unknown) => {
      patchMailAttachments(normalizeMailAttachments(item))
    })
    connection.on('MailAttachmentExported', (item: unknown) => {
      patchExportedAttachment(normalizeExportedMailAttachment(item))
    })
    connection.on('RulesUpdated', (items: OutlookRuleDto[]) => {
      rules.value = items
      loadingRules.value = false
      completeOperation()
    })
    connection.on('CategoriesUpdated', (items: unknown) => {
      categories.value = normalizeOutlookCategories(items)
      initialCategoriesFetchCompleted = true
      loadingCategories.value = false
      completeOperation()
    })
    connection.on('CalendarUpdated', (items: CalendarEventDto[]) => {
      calendarEvents.value = items
      loadingCalendar.value = false
      completeOperation()
    })
    connection.on('CommandResult', (result: OutlookCommandResult) => {
      const mailId = mailBodyCommandIds.get(result.commandId)
      if (mailId) completeMailBodyLoad(mailId)
      const attachmentCommand = attachmentCommandIds.get(result.commandId)
      if (attachmentCommand?.type === 'fetch') completeAttachmentLoad(attachmentCommand.mailId)
      if (attachmentCommand?.type === 'export' && attachmentCommand.attachmentId) {
        completeAttachmentExport(attachmentCommand.mailId, attachmentCommand.attachmentId)
      }
      completeOperation(result.commandId)
      if (activeMailSearchCommandId && result.commandId === activeMailSearchCommandId && !result.success) {
        loadingMailSearch.value = false
        activeMailSearchCommandId = ''
        window.clearTimeout(mailSearchTimeoutId)
      }
    })
    connection.on('NewChatMessage', async (message: ChatMessageDto) => {
      chatMessages.value = [...chatMessages.value, message]
      await scrollChatToBottom()
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
      void runStartupOutlookSync()
    } catch {
      signalRState.value = 'disconnected'
    }
  }

  watch(
    () => selectedMail.value?.id,
    () => resetMailPropertiesDraft(selectedMail.value),
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
  })

  onBeforeUnmount(() => {
    unmounted = true
    window.removeEventListener('click', closeFolderContextMenu)
    window.clearTimeout(operationTimeoutId)
    window.clearTimeout(mailSearchTimeoutId)
    clearMailBodyLoads()
    clearAttachmentLoads()
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
    htmlMailIndexes,
    loadingCalendar,
    loadingCategories,
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
    mailSearchResults,
    isAttachmentExporting,
    isAttachmentListLoading,
    isMailBodyLoading,
    mailHasBody,
    mailPropertiesDraft,
    mailPropertiesChanged,
    mailRange,
    mailStats,
    masterCategoryListExpanded,
    mails,
    moveDraggedMail,
    openFolderContextMenu,
    openExportedAttachment,
    openMailIndexes,
    operationLoading,
    outlookBusy,
    outlookBusyText,
    openCategoryManager,
    refreshAdminData,
    requestCalendar,
    requestCategories,
    requestFolders,
    requestSignalRPing,
    requestMails,
    requestMailSearch,
    resetMailPropertiesDraft,
    resetAttachmentExportRoot,
    removeMailCategoryDraft,
    saveAttachmentExportSettings,
    savingAttachmentExportSettings,
    fetchedMailFolderName,
    mailListNeedsFetch,
    selectedFolderName,
    selectedCalendarEvent,
    selectedFolderPath,
    selectedMail,
    selectedMailCategories,
    selectedMailAttachments,
    selectedMailFolderName,
    selectedMailHasIdentity,
    selectedMailHtml,
    selectedMailIndex,
    selectedMailIsOpen,
    selectFolder,
    selectCalendarEvent,
    selectMail,
    sendChat,
    showFolderMails,
    goToCurrentCalendarMonth,
    setDragOverFolder,
    setMailFlagDraft,
    signalRState,
    splitCategories,
    startMailDrag,
    switchView,
    toggleFolder,
    toggleMail,
    toggleMailFormat,
    updateCategoryColor,
    visibleFolders,
    visibleMasterCategories,
    hiddenMasterCategoryCount,
    toggleMasterCategoryList,
    flagEditorVisible,
  }
}
