import type { Ref } from 'vue'
import { ElMessage } from 'element-plus'
import { outlookApi } from '../api/outlook'
import type { MailItemDto } from '../models/outlook'
import { isMailSelectableFolder } from '../utils/folders'

type MailListControllerOptions = {
  folderMails: Ref<MailItemDto[]>
  mailSearchResults: Ref<MailItemDto[]>
  mailListMode: Ref<'folder' | 'search'>
  selectedMailIds: Ref<Set<string>>
  selectedMailIndex: Ref<number | null>
  draggedMailId: Ref<string>
  dragOverFolderPath: Ref<string>
  outlookBusy: Ref<boolean>
  folderOptions: Ref<any[]>
  mails: Ref<MailItemDto[]>
  manualOutlookDeleteMessage: string
  deletedFolderForPath: (path: string) => { folderPath: string; folderType: string; label: string } | null
  isInDeletedFolder: (path: string) => boolean
  loadCachedFolders: () => Promise<void>
  loadCachedMails: () => Promise<void>
  loadCachedMailSearchResults: () => Promise<void>
  runMailOperation: (action: () => Promise<unknown>, afterSuccess?: () => Promise<void>) => Promise<boolean>
}

export function useOutlookMailListController(options: MailListControllerOptions) {
  const {
    deletedFolderForPath,
    draggedMailId,
    dragOverFolderPath,
    folderMails,
    folderOptions,
    isInDeletedFolder,
    loadCachedFolders,
    loadCachedMailSearchResults,
    loadCachedMails,
    mailListMode,
    mailSearchResults,
    mails,
    manualOutlookDeleteMessage,
    outlookBusy,
    runMailOperation,
    selectedMailIds,
    selectedMailIndex,
  } = options
  let lastSelectedMailIndex = -1

  function pruneSelectedMailIds(items = mails.value) {
    const visibleIds = new Set(items.map((mail) => mail.id).filter(Boolean))
    selectedMailIds.value = new Set([...selectedMailIds.value].filter((id) => visibleIds.has(id)))
    if (lastSelectedMailIndex >= items.length) lastSelectedMailIndex = -1
  }

  function selectedBulkMoveMails() {
    return mails.value.filter((mail) => mail.id && selectedMailIds.value.has(mail.id))
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

  function clearSelectedMailIndex() {
    selectedMailIndex.value = null
    lastSelectedMailIndex = -1
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
      () => outlookApi.requestMoveMail({
        mailId: mail.id,
        sourceFolderPath: mail.folderPath,
        destinationFolderPath,
      }),
      async () => { await loadCachedFolders() },
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
      () => outlookApi.requestMoveMails({
        mailIds: selected.map((mail) => mail.id),
        sourceFolderPath: sourceFolderPaths.length === 1 ? sourceFolderPaths[0] : '',
        sourceFolderPaths,
        destinationFolderPath,
        continueOnError: true,
      }),
      async () => { await loadCachedFolders() },
    )
    if (succeeded) clearSelectedMails()
    else restoreFailedMailMove(snapshot)
  }

  function currentDragMailCount(mailId: string) {
    return selectedMailIds.value.has(mailId) && selectedMailIds.value.size > 1 ? selectedMailIds.value.size : 1
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
      () => outlookApi.requestDeleteMail({ mailId: mail.id, folderPath: mail.folderPath }),
      async () => { await Promise.allSettled([loadCachedMails(), loadCachedMailSearchResults(), loadCachedFolders()]) },
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

  return {
    clearMailDrag,
    clearSelectedMailIndex,
    clearSelectedMails,
    deleteMail,
    moveDraggedMail,
    pruneSelectedMailIds,
    selectMail,
    selectOnlyMail,
    setDragOverFolder,
    startMailDrag,
  }
}
