import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import * as signalR from '@microsoft/signalr'
import { normalizeMailBody, normalizeMailItem, normalizeMailItems, normalizeOutlookCategories, outlookApi } from '../api/outlook'
import type {
  AddinLogEntry,
  AddinStatusDto,
  AppView,
  CalendarEventDto,
  ChatMessageDto,
  FolderSyncBatchDto,
  FolderSyncBeginDto,
  FolderSyncCompleteDto,
  FolderTreeNode,
  MailBodyDto,
  MailItemDto,
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
import { applyFolderBatch, buildFolderTree, collectFolderOptions, folderType, visibleRootFolders } from '../utils/folders'
import {
  addMonths,
  dateInputToIso,
  defaultFlagRequest,
  flagIntervalLabel,
  flagIntervalOptions,
  isDefaultFlagRequest,
  mergeStores,
  monthEndExclusive,
  monthStart,
  sleep,
  splitCategories,
  toDateInput,
  toDateKey,
  todayInputValue,
} from '../utils/outlookDashboardHelpers'

export function useOutlookDashboard() {
  const activeView = ref<AppView>('outlook')
  const signalRState = ref<SignalRState>('disconnected')
  const folders = ref<FolderTreeNode[]>([])
  const folderStores = ref<OutlookStoreDto[]>([])
  const mails = ref<MailItemDto[]>([])
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
  const selectedFolderPath = ref('')
  const fetchedMailFolderPath = ref('')
  const pendingMailFolderPath = ref('')
  const selectedMailIndex = ref<number | null>(null)
  const selectedMailHtml = ref(false)
  const activePropertyLibrarySections = ref(['property-library'])
  const activeMailPropertySections = ref(['set-mail-properties'])
  const expandedFolders = ref<Set<string>>(new Set())
  const openMailIndexes = ref<Set<number>>(new Set())
  const htmlMailIndexes = ref<Set<number>>(new Set())
  const loadingMailBodyIds = ref<Set<string>>(new Set())
  const mailRange = ref('1d')
  const mailCount = ref(10)
  const chatText = ref('')
  const loadingFolders = ref(true)
  const loadingMails = ref(true)
  const loadingRules = ref(false)
  const loadingCategories = ref(true)
  const loadingCalendar = ref(false)
  const loadingSignalRPing = ref(false)
  const operationLoading = ref(false)
  const mailPropertiesDraft = ref({
    isRead: false,
    flagInterval: 'none',
    flagRequest: '',
    taskStartDate: '',
    taskDueDate: '',
    taskCompletedDate: '',
    categories: [] as string[],
  })
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
  const mailBodyCommandIds = new Map<string, string>()
  const mailBodyTimeoutIds = new Map<string, number>()

  const visibleFolders = computed(() => visibleRootFolders(folders.value))

  const mailStats = computed(() => ({
    unread: mails.value.filter((mail) => !mail.isRead).length,
    flagged: mails.value.filter((mail) => mail.isMarkedAsTask).length,
    highImportance: mails.value.filter((mail) => mail.importance === 'high').length,
    categorized: mails.value.filter((mail) => Boolean(mail.categories)).length,
  }))

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

  const calendarEventsByDay = computed(() => {
    const groups = new Map<string, CalendarEventDto[]>()
    for (const event of calendarEvents.value) {
      const key = toDateKey(new Date(event.start))
      const items = groups.get(key) ?? []
      items.push(event)
      groups.set(key, items)
    }

    for (const items of groups.values()) {
      items.sort((a, b) => new Date(a.start).getTime() - new Date(b.start).getTime())
    }

    return groups
  })

  const calendarWeeks = computed(() => {
    const first = monthStart(calendarMonthDate.value)
    const gridStart = new Date(first)
    gridStart.setDate(first.getDate() - first.getDay())
    const todayKey = toDateKey(new Date())

    return Array.from({ length: 6 }, (_, weekIndex) =>
      Array.from({ length: 7 }, (_, dayIndex) => {
        const date = new Date(gridStart)
        date.setDate(gridStart.getDate() + weekIndex * 7 + dayIndex)
        const key = toDateKey(date)
        return {
          key,
          date,
          dayNumber: date.getDate(),
          inMonth: date.getMonth() === calendarMonthDate.value.getMonth(),
          isToday: key === todayKey,
          events: calendarEventsByDay.value.get(key) ?? [],
        }
      }),
    )
  })

  const selectedFolderName = computed(() => {
    return folderNameForPath(selectedFolderPath.value)
  })

  const fetchedMailFolderName = computed(() => {
    return fetchedMailFolderPath.value ? folderNameForPath(fetchedMailFolderPath.value) : '尚未抓取郵件'
  })

  const selectedMailFolderName = computed(() => {
    return selectedMail.value?.folderPath ? folderNameForPath(selectedMail.value.folderPath) : '未選擇'
  })

  const mailListNeedsFetch = computed(() => {
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

  function folderNameForPath(path: string) {
    if (!path) return '未選擇'
    return folderOptions.value.find((folder) => folder.folderPath === path)?.label.trim() ?? path
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

  function resetMailPropertiesDraft(mail: MailItemDto | null) {
    if (!mail) return
    const flagInterval = mail.flagInterval || (mail.isMarkedAsTask ? 'today' : 'none')
    mailPropertiesDraft.value = {
      isRead: mail.isRead,
      flagInterval,
      flagRequest: isDefaultFlagRequest(mail.flagRequest) ? defaultFlagRequest(flagInterval) : mail.flagRequest,
      taskStartDate: toDateInput(mail.taskStartDate),
      taskDueDate: toDateInput(mail.taskDueDate),
      taskCompletedDate: toDateInput(mail.taskCompletedDate),
      categories: splitCategories(mail.categories),
    }
  }

  function setMails(items: MailItemDto[], preferredMailId = selectedMail.value?.id ?? '') {
    const wasOpen = selectedMailIndex.value !== null && openMailIndexes.value.has(selectedMailIndex.value)
    clearMailBodyLoads()
    mails.value = items

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
    const index = mails.value.findIndex((mail) => mail.id === nextMail.id)
    if (index < 0) return
    const items = [...mails.value]
    items[index] = {
      ...nextMail,
      body: nextMail.body || items[index].body,
      bodyHtml: nextMail.bodyHtml || items[index].bodyHtml,
    }
    mails.value = items
  }

  function patchMailBody(body: MailBodyDto) {
    if (!body.mailId) return
    const index = mails.value.findIndex((mail) => mail.id === body.mailId)
    if (index < 0) return
    const items = [...mails.value]
    items[index] = {
      ...items[index],
      body: body.body,
      bodyHtml: body.bodyHtml,
      folderPath: body.folderPath || items[index].folderPath,
    }
    mails.value = items
    completeMailBodyLoad(body.mailId)
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
      const result = await outlookApi.requestFolders()
      if (result.status === 'mocked') {
        await loadCachedFolders()
        initialFoldersFetchCompleted = true
        loadingFolders.value = false
        operationLoading.value = false
      }
    } catch {
      initialFoldersFetchCompleted = true
      loadingFolders.value = false
    }
  }

  function folderStore(folder: FolderTreeNode) {
    return folderStores.value.find((store) => store.storeId === folder.storeId)
  }

  function findPreferredInboxFolder() {
    const inboxes = folderOptions.value.filter((folder) => folderType(folder.name) === 'inbox')
    return (
      inboxes.find((folder) => folderStore(folder)?.storeKind?.toLowerCase() === 'ost')
      ?? inboxes.find((folder) => ['exchange', 'ost'].includes(folderStore(folder)?.storeKind?.toLowerCase() ?? ''))
      ?? inboxes[0]
      ?? folderOptions.value[0]
      ?? null
    )
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

  function toggleFolder(path: string) {
    if (outlookBusy.value) return
    const next = new Set(expandedFolders.value)
    if (next.has(path)) next.delete(path)
    else next.add(path)
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
    } catch {
      pendingMailFolderPath.value = ''
      initialMailsFetchCompleted = true
      loadingMails.value = false
    }
  }

  async function requestRules() {
    if (outlookBusy.value) return
    loadingRules.value = true
    try {
      await outlookApi.requestRules()
    } catch {
      loadingRules.value = false
    }
  }

  async function requestCategories(force = false) {
    if (outlookBusy.value && !force) return
    loadingCategories.value = true
    try {
      const result = await outlookApi.requestCategories()
      if (result.status === 'mocked') {
        await loadCachedCategories()
        initialCategoriesFetchCompleted = true
        loadingCategories.value = false
        operationLoading.value = false
      }
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

  async function waitForFoldersReady(timeoutMs = 12000) {
    const started = Date.now()
    while (!unmounted && (loadingFolders.value || folderOptions.value.length === 0) && Date.now() - started < timeoutMs) {
      await sleep(100)
    }
  }

  async function waitForInitialFetch(done: () => boolean, timeoutMs = 12000) {
    const started = Date.now()
    while (!unmounted && !done() && Date.now() - started < timeoutMs) {
      await sleep(100)
    }
  }

  async function waitForNotificationSignalRConnected(timeoutMs = 12000) {
    const started = Date.now()
    while (!unmounted && signalRState.value !== 'connected' && Date.now() - started < timeoutMs) {
      await sleep(100)
    }
    return signalRState.value === 'connected'
  }

  async function runStartupOutlookSync() {
    if (startupSyncStarted) return
    const connected = await waitForNotificationSignalRConnected()
    if (!connected || unmounted) return
    startupSyncStarted = true
    await sleep(500)
    if (unmounted) return
    await requestFolders(true)
    await waitForFoldersReady()
    if (!initialFoldersFetchCompleted) {
      initialFoldersFetchCompleted = true
      loadingFolders.value = false
    }
    if (unmounted) return
    selectInboxFolder()

    await sleep(500)
    if (unmounted) return
    await requestCategories(true)
    await waitForInitialFetch(() => initialCategoriesFetchCompleted || !loadingCategories.value)

    await sleep(500)
    if (unmounted) return
    if (!selectedFolderPath.value) selectInboxFolder()
    mailRange.value = '1d'
    mailCount.value = 10
    await requestMails(true)
    await waitForInitialFetch(() => initialMailsFetchCompleted || !loadingMails.value)
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

  function isMailBodyLoading(mail: MailItemDto) {
    return Boolean(mail.id && loadingMailBodyIds.value.has(mail.id))
  }

  function mailHasBody(mail: MailItemDto) {
    return Boolean(mail.body || mail.bodyHtml)
  }

  async function requestMailBody(mail: MailItemDto) {
    if (!mail.id?.trim() || mailHasBody(mail) || isMailBodyLoading(mail)) return
    loadingMailBodyIds.value = new Set(loadingMailBodyIds.value).add(mail.id)
    try {
      const response = await outlookApi.requestMailBody({
        mailId: mail.id,
        folderPath: mail.folderPath,
      })
      mailBodyCommandIds.set(response.commandId, mail.id)
      const timeoutId = window.setTimeout(() => completeMailBodyLoad(mail.id), 30000)
      mailBodyTimeoutIds.set(mail.id, timeoutId)
    } catch {
      completeMailBodyLoad(mail.id)
    }
  }

  async function applyMailProperties(mail: MailItemDto) {
    if (!mail.id?.trim()) return
    const selectedCategories = [...new Set(mailPropertiesDraft.value.categories.map((category) => category.trim()).filter(Boolean))]
    const existingCategoryNames = new Set(categories.value.map((category) => category.name.toLowerCase()))
    const newCategories = selectedCategories
      .filter((category) => !existingCategoryNames.has(category.toLowerCase()))
      .map((name) => ({ name, color: 'olCategoryColorNone', colorValue: 0, shortcutKey: '' }))
    const isCustomFlag = mailPropertiesDraft.value.flagInterval === 'custom'
    const body: MailPropertiesCommandRequest = {
      mailId: mail.id,
      folderPath: mail.folderPath,
      isRead: mailPropertiesDraft.value.isRead,
      flagInterval: mailPropertiesDraft.value.flagInterval,
      flagRequest: mailPropertiesDraft.value.flagRequest || defaultFlagRequest(mailPropertiesDraft.value.flagInterval),
      taskStartDate: isCustomFlag ? dateInputToIso(mailPropertiesDraft.value.taskStartDate) : undefined,
      taskDueDate: isCustomFlag ? dateInputToIso(mailPropertiesDraft.value.taskDueDate) : undefined,
      taskCompletedDate: mailPropertiesDraft.value.flagInterval === 'complete' ? dateInputToIso(mailPropertiesDraft.value.taskCompletedDate) : undefined,
      categories: selectedCategories,
      newCategories,
    }
    await runMailOperation(() => outlookApi.requestUpdateMailProperties(body))
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

  async function switchView(view: AppView) {
    activeView.value = view
    if (view === 'admin') await refreshAdminData()
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
    connection.on('FolderSyncStarted', (info: FolderSyncBeginDto) => {
      if (info.reset) {
        folders.value = []
        folderStores.value = []
      }
      loadingFolders.value = true
    })
    connection.on('FoldersPatched', (batch: FolderSyncBatchDto) => {
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
    connection.on('MailUpdated', (item: unknown) => {
      const mail = normalizeMailItem(item)
      patchMail(mail)
      if (mailHasBody(mail)) completeMailBodyLoad(mail.id)
      completeOperation()
    })
    connection.on('MailBodyUpdated', (item: unknown) => {
      patchMailBody(normalizeMailBody(item))
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
      completeOperation(result.commandId)
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
      loadCachedRules(),
      loadCachedCategories(),
      loadCachedCalendar(),
      loadChat(),
      refreshAdminData(),
    ])
  })

  onBeforeUnmount(() => {
    unmounted = true
    window.removeEventListener('click', closeFolderContextMenu)
    window.clearTimeout(operationTimeoutId)
    clearMailBodyLoads()
    void connection?.stop()
  })

  return {
    activeView,
    activeMailPropertySections,
    activePropertyLibrarySections,
    addCategoryToMasterList,
    addinLogs,
    addinStatus,
    applyMailProperties,
    calendarEvents,
    calendarMonthLabel,
    calendarWeekdays,
    calendarWeeks,
    cancelCreateFolder,
    categories,
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
    dragOverFolderPath,
    draggedMailId,
    expandedFolders,
    fetchMailsFromContext,
    flagIntervalLabel,
    flagIntervalOptions,
    folderContextMenu,
    folderStores,
    htmlMailIndexes,
    loadingCalendar,
    loadingCategories,
    loadingFolders,
    loadingMails,
    loadingSignalRPing,
    mailCount,
    mailHtmlSandbox,
    isMailBodyLoading,
    mailHasBody,
    mailPropertiesDraft,
    mailRange,
    mailStats,
    mails,
    moveDraggedMail,
    openFolderContextMenu,
    openMailIndexes,
    operationLoading,
    outlookBusy,
    outlookBusyText,
    refreshAdminData,
    requestCalendar,
    requestCategories,
    requestFolders,
    requestSignalRPing,
    requestMails,
    resetMailPropertiesDraft,
    fetchedMailFolderName,
    mailListNeedsFetch,
    selectedFolderName,
    selectedCalendarEvent,
    selectedFolderPath,
    selectedMail,
    selectedMailCategories,
    selectedMailFolderName,
    selectedMailHasIdentity,
    selectedMailHtml,
    selectedMailIndex,
    selectedMailIsOpen,
    selectFolder,
    selectCalendarEvent,
    selectMail,
    sendChat,
    goToCurrentCalendarMonth,
    setDragOverFolder,
    signalRState,
    splitCategories,
    startMailDrag,
    switchView,
    toggleFolder,
    toggleMail,
    toggleMailFormat,
    updateCategoryColor,
    visibleFolders,
  }
}
