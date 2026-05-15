import type { Ref } from 'vue'
import { ElMessage } from 'element-plus'
import { outlookApi } from '../api/outlook'
import type { FolderTreeNode, MailItemDto } from '../models/outlook'
import { cloneTree, isMailSelectableFolder } from '../utils/folders'
import { canMoveOutlookItem } from '../utils/outlookItemTypes'

type MailListControllerOptions = {
  folderMails: Ref<MailItemDto[]>
  folders: Ref<FolderTreeNode[]>
  mailSearchResults: Ref<MailItemDto[]>
  mailListMode: Ref<'folder' | 'search'>
  mailDragPreview: Ref<{ visible: boolean; x: number; y: number; subject: string; count: number }>
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
  runMailOperation: (action: () => Promise<unknown>, afterSuccess?: (response?: unknown) => Promise<void>) => Promise<boolean>
}

export function useOutlookMailListController(options: MailListControllerOptions) {
  const {
    deletedFolderForPath,
    draggedMailId,
    dragOverFolderPath,
    folderMails,
    folders,
    folderOptions,
    isInDeletedFolder,
    mailListMode,
    mailDragPreview,
    mailSearchResults,
    mails,
    manualOutlookDeleteMessage,
    outlookBusy,
    runMailOperation,
    selectedMailIds,
    selectedMailIndex,
  } = options
  let lastSelectedMailIndex = -1
  let pointerDrag:
    | {
      mail: MailItemDto
      index: number
      pointerId: number
      startX: number
      startY: number
      active: boolean
    }
    | null = null

  function pruneSelectedMailIds(items = mails.value) {
    const visibleIds = new Set(items.map((mail) => mail.id).filter(Boolean))
    selectedMailIds.value = new Set([...selectedMailIds.value].filter((id) => visibleIds.has(id)))
    if (lastSelectedMailIndex >= items.length) lastSelectedMailIndex = -1
  }

  function selectedBulkMoveMails() {
    return mails.value.filter((mail) => mail.id && selectedMailIds.value.has(mail.id))
  }

  function selectedBulkDeleteMails() {
    return selectedBulkMoveMails().filter((mail) => canMoveOutlookItem(mail) && !isInDeletedFolder(mail.folderPath))
  }

  function shouldBulkDeleteFromRow(mail: MailItemDto) {
    return Boolean(mail.id && selectedMailIds.value.size > 1 && selectedMailIds.value.has(mail.id))
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
      folders: cloneTree(folders.value),
      mailSearchResults: [...mailSearchResults.value],
      selectedMailIds: new Set(selectedMailIds.value),
      selectedMailIndex: selectedMailIndex.value,
      lastSelectedMailIndex,
    }
  }

  function restoreMailListSnapshot(snapshot: ReturnType<typeof captureMailListSnapshot>) {
    folderMails.value = snapshot.folderMails
    folders.value = cloneTree(snapshot.folders)
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

  function hideDeletedMails(mailIds: string[]) {
    const deletedIds = new Set(mailIds.filter(Boolean))
    if (deletedIds.size === 0) return

    folderMails.value = folderMails.value.filter((mail) => !deletedIds.has(mail.id))
    mailSearchResults.value = mailSearchResults.value.filter((mail) => !deletedIds.has(mail.id))
    selectedMailIds.value = new Set([...selectedMailIds.value].filter((id) => !deletedIds.has(id)))
    const nextIndex = firstSelectedMailIndex(selectedMailIds.value)
    selectedMailIndex.value = nextIndex >= 0 ? nextIndex : null
    if (selectedMailIndex.value === null) lastSelectedMailIndex = -1
  }

  function restoreFailedMailMove(snapshot: ReturnType<typeof captureMailListSnapshot>) {
    restoreMailListSnapshot(snapshot)
    ElMessage.error('移動郵件失敗，已還原畫面。')
  }

  function applyFolderCountDeltas(deltas: Map<string, number>) {
    if (deltas.size === 0) return

    function patch(nodes: FolderTreeNode[]): FolderTreeNode[] {
      return nodes.map((folder) => ({
        ...folder,
        itemCount: Math.max(0, folder.itemCount + (deltas.get(folder.folderPath) ?? 0)),
        subFolders: patch(folder.subFolders),
      }))
    }

    folders.value = patch(folders.value)
  }

  function addFolderDelta(deltas: Map<string, number>, folderPath: string, delta: number) {
    if (!folderPath || delta === 0) return
    deltas.set(folderPath, (deltas.get(folderPath) ?? 0) + delta)
  }

  function moveFolderCountDeltas(items: MailItemDto[], destinationFolderPath: string) {
    const deltas = new Map<string, number>()
    for (const mail of items) addFolderDelta(deltas, mail.folderPath, -1)
    addFolderDelta(deltas, destinationFolderPath, items.length)
    return deltas
  }

  function deleteFolderCountDeltas(items: MailItemDto[]) {
    const deltas = new Map<string, number>()
    for (const mail of items) {
      addFolderDelta(deltas, mail.folderPath, -1)
      const deletedFolderPath = deletedFolderForPath(mail.folderPath)?.folderPath
      if (deletedFolderPath) addFolderDelta(deltas, deletedFolderPath, 1)
    }
    return deltas
  }

  function applyConfirmedFolderDeltas(deltas: Map<string, number>) {
    applyFolderCountDeltas(deltas)
  }

  async function moveMailToFolder(mail: MailItemDto, destinationFolderPath: string) {
    if (!mail.id?.trim() || !destinationFolderPath || destinationFolderPath === mail.folderPath) return
    const snapshot = captureMailListSnapshot()
    const folderDeltas = moveFolderCountDeltas([mail], destinationFolderPath)
    hideMovedMails([mail.id])
    const succeeded = await runMailOperation(
      () => outlookApi.requestMoveMail({
        mailId: mail.id,
        sourceFolderPath: mail.folderPath,
        destinationFolderPath,
      }),
      async () => { applyConfirmedFolderDeltas(folderDeltas) },
    )
    if (!succeeded) restoreFailedMailMove(snapshot)
  }

  async function moveSelectedMailsToFolder(destinationFolderPath: string) {
    const selected = selectedBulkMoveMails()
    if (selected.length === 0 || !destinationFolderPath || outlookBusy.value) return
    const sourceFolderPaths = [...new Set(selected.map((mail) => mail.folderPath).filter(Boolean))]
    const snapshot = captureMailListSnapshot()
    const folderDeltas = moveFolderCountDeltas(selected, destinationFolderPath)
    hideMovedMails(selected.map((mail) => mail.id))
    const succeeded = await runMailOperation(
      () => outlookApi.requestMoveMails({
        mailIds: selected.map((mail) => mail.id),
        sourceFolderPath: sourceFolderPaths.length === 1 ? sourceFolderPaths[0] : '',
        sourceFolderPaths,
        destinationFolderPath,
        continueOnError: true,
      }),
      async () => { applyConfirmedFolderDeltas(folderDeltas) },
    )
    if (succeeded) clearSelectedMails()
    else restoreFailedMailMove(snapshot)
  }

  async function deleteMail(mail: MailItemDto) {
    if (!mail?.id?.trim() || outlookBusy.value) return
    if (shouldBulkDeleteFromRow(mail)) {
      await deleteSelectedMails()
      return
    }
    if (isInDeletedFolder(mail.folderPath)) {
      ElMessage.warning(manualOutlookDeleteMessage)
      return
    }
    const deletedFolder = deletedFolderForPath(mail.folderPath) ?? folderOptions.value.find((folder) => folder.folderType === 'Deleted')
    const targetName = deletedFolder?.label.trim() || '刪除的郵件 / Deleted Items'
    const confirmed = window.confirm(`將郵件「${mail.subject || mail.id}」移到「${targetName}」？`)
    if (!confirmed) return
    const snapshot = captureMailListSnapshot()
    const folderDeltas = deleteFolderCountDeltas([mail])
    hideDeletedMails([mail.id])
    const succeeded = await runMailOperation(
      () => outlookApi.requestDeleteMail({ mailId: mail.id, folderPath: mail.folderPath }),
      async () => { applyConfirmedFolderDeltas(folderDeltas) },
    )
    if (!succeeded) restoreFailedMailMove(snapshot)
  }

  async function deleteSelectedMails() {
    const selected = selectedBulkDeleteMails()
    if (selected.length === 0 || outlookBusy.value) {
      if (selectedMailIds.value.size > 0) ElMessage.warning(manualOutlookDeleteMessage)
      return
    }

    const targetNames = [...new Set(selected
      .map((mail) => deletedFolderForPath(mail.folderPath)?.label.trim())
      .filter(Boolean))]
    const targetName = targetNames.length === 1 ? targetNames[0] : '各自的刪除資料夾 / Deleted Items'
    const confirmed = window.confirm(`將選取的 ${selected.length} 封郵件移到「${targetName}」？`)
    if (!confirmed) return

    const snapshot = captureMailListSnapshot()
    const folderDeltas = deleteFolderCountDeltas(selected)
    hideDeletedMails(selected.map((mail) => mail.id))

    for (const [index, mail] of selected.entries()) {
      const succeeded = await runMailOperation(
        () => outlookApi.requestDeleteMail({ mailId: mail.id, folderPath: mail.folderPath }),
        index === selected.length - 1 ? async () => { applyConfirmedFolderDeltas(folderDeltas) } : undefined,
      )
      if (!succeeded) {
        restoreFailedMailMove(snapshot)
        return
      }
    }

    clearSelectedMails()
  }

  function clearMailDrag() {
    draggedMailId.value = ''
    dragOverFolderPath.value = ''
    mailDragPreview.value = { visible: false, x: 0, y: 0, subject: '', count: 0 }
    document.body.classList.remove('mail-dragging')
  }

  function folderPathAtPoint(x: number, y: number) {
    const target = document.elementFromPoint(x, y)?.closest<HTMLElement>('[data-mail-drop-folder-path]')
    return target?.dataset.mailDropFolderPath || ''
  }

  function beginPointerMailDrag() {
    if (!pointerDrag || pointerDrag.active) return
    pointerDrag.active = true
    if (!selectedMailIds.value.has(pointerDrag.mail.id)) selectOnlyMail(pointerDrag.index)
    draggedMailId.value = pointerDrag.mail.id
    const dragCount = Math.max(1, selectedMailIds.value.size)
    mailDragPreview.value = {
      visible: true,
      x: pointerDrag.startX,
      y: pointerDrag.startY,
      subject: dragCount > 1 ? `移動 ${dragCount} 封郵件` : pointerDrag.mail.subject || '(No subject)',
      count: dragCount,
    }
    document.body.classList.add('mail-dragging')
  }

  function handleMailPointerMove(event: PointerEvent) {
    if (!pointerDrag || event.pointerId !== pointerDrag.pointerId) return
    const moved = Math.hypot(event.clientX - pointerDrag.startX, event.clientY - pointerDrag.startY)
    if (!pointerDrag.active && moved < 6) return
    beginPointerMailDrag()
    event.preventDefault()
    mailDragPreview.value = { ...mailDragPreview.value, visible: true, x: event.clientX, y: event.clientY }
    setDragOverFolder(folderPathAtPoint(event.clientX, event.clientY))
  }

  async function handleMailPointerUp(event: PointerEvent) {
    if (!pointerDrag || event.pointerId !== pointerDrag.pointerId) return
    const wasActive = pointerDrag.active
    const dropFolderPath = wasActive ? folderPathAtPoint(event.clientX, event.clientY) : ''
    window.removeEventListener('pointermove', handleMailPointerMove)
    window.removeEventListener('pointerup', handleMailPointerUp)
    window.removeEventListener('pointercancel', handleMailPointerCancel)
    pointerDrag = null
    if (!wasActive) return
    event.preventDefault()
    if (dropFolderPath) await moveDraggedMail(dropFolderPath)
    else clearMailDrag()
  }

  function handleMailPointerCancel(event: PointerEvent) {
    if (!pointerDrag || event.pointerId !== pointerDrag.pointerId) return
    window.removeEventListener('pointermove', handleMailPointerMove)
    window.removeEventListener('pointerup', handleMailPointerUp)
    window.removeEventListener('pointercancel', handleMailPointerCancel)
    pointerDrag = null
    clearMailDrag()
  }

  function startMailPointerDrag(mail: MailItemDto, index: number, event: PointerEvent) {
    if (event.button !== 0 || !mail.id?.trim() || outlookBusy.value || !canMoveOutlookItem(mail)) return
    pointerDrag = {
      mail,
      index,
      pointerId: event.pointerId,
      startX: event.clientX,
      startY: event.clientY,
      active: false,
    }
    window.addEventListener('pointermove', handleMailPointerMove, { passive: false })
    window.addEventListener('pointerup', handleMailPointerUp)
    window.addEventListener('pointercancel', handleMailPointerCancel)
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
    deleteSelectedMails,
    moveDraggedMail,
    pruneSelectedMailIds,
    selectMail,
    selectOnlyMail,
    setDragOverFolder,
    startMailPointerDrag,
  }
}
