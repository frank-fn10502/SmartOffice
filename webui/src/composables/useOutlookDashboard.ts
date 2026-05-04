import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import * as signalR from '@microsoft/signalr'
import { normalizeCategoryColor, normalizeMailItems, normalizeOutlookCategories, outlookApi } from '../api/outlook'
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
  MailItemDto,
  MailPropertiesCommandRequest,
  OutlookStoreDto,
  OutlookCategoryDto,
  OutlookRuleDto,
  SignalRState,
} from '../models/outlook'
import { applyFolderBatch, buildFolderTree, collectFolderOptions, folderType, visibleRootFolders } from '../utils/folders'

const flagIntervalOptions = [
  { label: '不設定旗標', value: 'none' },
  { label: '今天', value: 'today' },
  { label: '明天', value: 'tomorrow' },
  { label: '本週', value: 'this_week' },
  { label: '下週', value: 'next_week' },
  { label: '無日期', value: 'no_date' },
  { label: '自訂日期', value: 'custom' },
  { label: '標示完成', value: 'complete' },
]

const categoryColorOptions = [
  { label: '無色', value: 'olCategoryColorNone', color: '#eef2f7' },
  { label: '紅色', value: 'olCategoryColorRed', color: '#f87171' },
  { label: '橘色', value: 'olCategoryColorOrange', color: '#fb923c' },
  { label: '桃色', value: 'olCategoryColorPeach', color: '#f9a8d4' },
  { label: '黃色', value: 'olCategoryColorYellow', color: '#facc15' },
  { label: '綠色', value: 'olCategoryColorGreen', color: '#22c55e' },
  { label: '青色', value: 'olCategoryColorTeal', color: '#14b8a6' },
  { label: '橄欖', value: 'olCategoryColorOlive', color: '#84cc16' },
  { label: '藍色', value: 'olCategoryColorBlue', color: '#38bdf8' },
  { label: '紫色', value: 'olCategoryColorPurple', color: '#a78bfa' },
  { label: '栗色', value: 'olCategoryColorMaroon', color: '#be123c' },
  { label: '鋼藍', value: 'olCategoryColorSteel', color: '#64748b' },
  { label: '深鋼藍', value: 'olCategoryColorDarkSteel', color: '#475569' },
  { label: '灰色', value: 'olCategoryColorGray', color: '#94a3b8' },
  { label: '深灰', value: 'olCategoryColorDarkGray', color: '#64748b' },
  { label: '黑色', value: 'olCategoryColorBlack', color: '#111827' },
  { label: '深紅', value: 'olCategoryColorDarkRed', color: '#b91c1c' },
  { label: '深橘', value: 'olCategoryColorDarkOrange', color: '#c2410c' },
  { label: '深桃色', value: 'olCategoryColorDarkPeach', color: '#db2777' },
  { label: '深黃', value: 'olCategoryColorDarkYellow', color: '#ca8a04' },
  { label: '深綠', value: 'olCategoryColorDarkGreen', color: '#15803d' },
  { label: '深青', value: 'olCategoryColorDarkTeal', color: '#0f766e' },
  { label: '深橄欖', value: 'olCategoryColorDarkOlive', color: '#4d7c0f' },
  { label: '深藍', value: 'olCategoryColorDarkBlue', color: '#2563eb' },
  { label: '深紫', value: 'olCategoryColorDarkPurple', color: '#7e22ce' },
  { label: '深栗色', value: 'olCategoryColorDarkMaroon', color: '#9f1239' },
]

function defaultFlagRequest(value: string) {
  return value === 'none' ? '' : flagIntervalLabel(value)
}

function flagIntervalLabel(value?: string) {
  return flagIntervalOptions.find((option) => option.value === value)?.label ?? '旗標'
}

function isDefaultFlagRequest(value: string, previousInterval = '') {
  const normalized = value.trim()
  return (
    !normalized
    || normalized === 'Follow up'
    || normalized === '旗標'
    || normalized === flagIntervalLabel(previousInterval)
    || flagIntervalOptions.some((option) => option.label === normalized)
  )
}

function toDateInput(value?: string) {
  if (!value) return ''
  if (/^\d{4}-\d{2}-\d{2}/.test(value)) return value.slice(0, 10)
  const date = new Date(value)
  if (Number.isNaN(date.getTime())) return ''
  return date.toISOString().slice(0, 10)
}

function dateInputToIso(value: string) {
  return value ? `${value}T00:00:00` : undefined
}

function todayInputValue() {
  const now = new Date()
  const year = now.getFullYear()
  const month = `${now.getMonth() + 1}`.padStart(2, '0')
  const day = `${now.getDate()}`.padStart(2, '0')
  return `${year}-${month}-${day}`
}

function sleep(ms: number) {
  return new Promise((resolve) => window.setTimeout(resolve, ms))
}

function toDateKey(date: Date) {
  const year = date.getFullYear()
  const month = `${date.getMonth() + 1}`.padStart(2, '0')
  const day = `${date.getDate()}`.padStart(2, '0')
  return `${year}-${month}-${day}`
}

function monthStart(date: Date) {
  return new Date(date.getFullYear(), date.getMonth(), 1)
}

function monthEndExclusive(date: Date) {
  return new Date(date.getFullYear(), date.getMonth() + 1, 1)
}

function addMonths(date: Date, count: number) {
  return new Date(date.getFullYear(), date.getMonth() + count, 1)
}

function splitCategories(value: string) {
  return value
    .split(',')
    .map((category) => category.trim())
    .filter(Boolean)
}

function categoryColorStyle(value?: string) {
  const colorValue = normalizeCategoryColor(value ?? '')
  const color = categoryColorOptions.find((option) => option.value === colorValue)?.color ?? categoryColorOptions[0].color
  return { backgroundColor: color }
}

function categoryTextColor(backgroundColor: string) {
  const hex = backgroundColor.replace('#', '')
  if (hex.length !== 6) return '#172033'
  const red = Number.parseInt(hex.slice(0, 2), 16)
  const green = Number.parseInt(hex.slice(2, 4), 16)
  const blue = Number.parseInt(hex.slice(4, 6), 16)
  const luminance = (red * 299 + green * 587 + blue * 114) / 1000
  return luminance > 150 ? '#172033' : '#ffffff'
}

function categoryOptionColor(value?: string) {
  const colorValue = normalizeCategoryColor(value ?? '')
  return categoryColorOptions.find((option) => option.value === colorValue)?.color ?? categoryColorOptions[0].color
}

function mergeStores(current: OutlookStoreDto[], incoming: OutlookStoreDto[]) {
  const stores = [...current]
  for (const next of incoming) {
    const index = stores.findIndex((store) => store.storeId === next.storeId)
    if (index < 0) stores.push(next)
    else stores[index] = next
  }
  return stores
}

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
  const selectedMailIndex = ref<number | null>(null)
  const selectedMailHtml = ref(false)
  const activePropertyLibrarySections = ref(['property-library'])
  const activeMailPropertySections = ref(['set-mail-properties'])
  const expandedFolders = ref<Set<string>>(new Set())
  const openMailIndexes = ref<Set<number>>(new Set())
  const htmlMailIndexes = ref<Set<number>>(new Set())
  const mailRange = ref('1d')
  const mailCount = ref(10)
  const chatText = ref('')
  const loadingFolders = ref(false)
  const loadingMails = ref(false)
  const loadingRules = ref(false)
  const loadingCategories = ref(false)
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
    return folderOptions.value.find((folder) => folder.folderPath === selectedFolderPath.value)?.label.trim() ?? '未選擇'
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

  async function loadCachedFolders() {
    const snapshot = await outlookApi.getFolders()
    folderStores.value = snapshot.stores
    folders.value = buildFolderTree(snapshot)
    selectDefaultFolder()
  }

  async function requestFolders() {
    if (outlookBusy.value) return
    loadingFolders.value = true
    try {
      const result = await outlookApi.requestFolders()
      if (result.status === 'mocked') {
        await loadCachedFolders()
        loadingFolders.value = false
        operationLoading.value = false
      }
    } catch {
      loadingFolders.value = false
    }
  }

  function selectDefaultFolder() {
    if (selectedFolderPath.value || visibleFolders.value.length === 0) return
    const inbox = visibleFolders.value.find((folder) => folderType(folder.name) === 'inbox')
    selectedFolderPath.value = inbox?.folderPath ?? visibleFolders.value[0]?.folderPath ?? ''
  }

  function selectInboxFolder() {
    const inbox = folderOptions.value.find((folder) => folderType(folder.name) === 'inbox')
    selectedFolderPath.value = inbox?.folderPath ?? folderOptions.value[0]?.folderPath ?? ''
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
    setMails(await outlookApi.getMails())
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

  async function requestMails() {
    if (outlookBusy.value) return
    loadingMails.value = true
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

  async function requestCategories() {
    if (outlookBusy.value) return
    loadingCategories.value = true
    try {
      await outlookApi.requestCategories()
    } catch {
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

  async function waitForOutlookIdle(timeoutMs = 12000) {
    const started = Date.now()
    while (!unmounted && outlookBusy.value && Date.now() - started < timeoutMs) {
      await sleep(100)
    }
  }

  async function runStartupOutlookSync() {
    await sleep(1500)
    if (unmounted) return
    await waitForOutlookIdle()
    await requestFolders()

    await waitForOutlookIdle()
    await sleep(500)
    if (unmounted) return
    await requestCategories()

    await waitForOutlookIdle()
    await sleep(500)
    if (unmounted) return
    selectInboxFolder()
    mailRange.value = '1d'
    mailCount.value = 10
    await requestMails()
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
    operationLoading.value = true
    try {
      await action()
    } catch {
      operationLoading.value = false
    }
  }

  async function applyMailProperties(mail: MailItemDto) {
    if (!mail.id?.trim()) return
    const selectedCategories = [...new Set(mailPropertiesDraft.value.categories.map((category) => category.trim()).filter(Boolean))]
    const existingCategoryNames = new Set(categories.value.map((category) => category.name.toLowerCase()))
    const newCategories = selectedCategories
      .filter((category) => !existingCategoryNames.has(category.toLowerCase()))
      .map((name) => ({ name, color: 'olCategoryColorNone', shortcutKey: '' }))
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
      else next.add(index)
      openMailIndexes.value = next
      return
    }

    selectedMailIndex.value = index
    selectedMailHtml.value = false
    const next = new Set<number>()
    next.add(index)
    openMailIndexes.value = next
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
      if (batch.isFinal) operationLoading.value = false
    })
    connection.on('FolderSyncCompleted', (_info: FolderSyncCompleteDto) => {
      loadingFolders.value = false
      operationLoading.value = false
    })
    connection.on('MailsUpdated', (items: unknown) => {
      setMails(normalizeMailItems(items))
      loadingMails.value = false
      operationLoading.value = false
    })
    connection.on('RulesUpdated', (items: OutlookRuleDto[]) => {
      rules.value = items
      loadingRules.value = false
      operationLoading.value = false
    })
    connection.on('CategoriesUpdated', (items: unknown) => {
      categories.value = normalizeOutlookCategories(items)
      loadingCategories.value = false
      operationLoading.value = false
    })
    connection.on('CalendarUpdated', (items: CalendarEventDto[]) => {
      calendarEvents.value = items
      loadingCalendar.value = false
      operationLoading.value = false
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
    void runStartupOutlookSync()
  })

  onBeforeUnmount(() => {
    unmounted = true
    window.removeEventListener('click', closeFolderContextMenu)
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
    selectedFolderName,
    selectedCalendarEvent,
    selectedFolderPath,
    selectedMail,
    selectedMailCategories,
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
