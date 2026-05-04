import type { FolderDto, FolderSnapshotDto, FolderSyncBatchDto, FolderTreeNode, OutlookStoreDto } from '../models/outlook'

const hiddenFolderNames = [
  'common views',
  'finder',
  'reminders',
  'quick step',
  'conversation history',
  'conversation action',
  'server failures',
  'local failures',
  'conflicts',
  'sync issues',
  'rss',
  'social network',
  'people',
  'externalcontacts',
  'yammer',
]

export function isHiddenFolder(name: string) {
  const lowerName = name.toLowerCase()
  return hiddenFolderNames.some((hidden) => lowerName.includes(hidden))
}

export function visibleChildren(folder: FolderTreeNode) {
  return (folder.subFolders ?? []).filter((child) => !isHiddenFolder(child.name))
}

export function visibleRootFolders(folders: FolderTreeNode[]) {
  return folders.filter((root) => !isHiddenFolder(root.name))
}

export function collectFolderOptions(items: FolderTreeNode[], level = 0): Array<FolderTreeNode & { label: string }> {
  return items.flatMap((folder) => {
    const indent = level === 0 ? '' : `${'　'.repeat(level)}`
    return [
      { ...folder, label: `${indent}${folder.name}` },
      ...collectFolderOptions(visibleChildren(folder), level + 1),
    ]
  })
}

export function findFolderByPath(items: FolderTreeNode[], path: string): FolderTreeNode | null {
  for (const folder of items) {
    if (folder.folderPath === path) return folder
    const child = findFolderByPath(visibleChildren(folder), path)
    if (child) return child
  }
  return null
}

export function folderType(name: string) {
  const lowerName = name.toLowerCase()
  if (lowerName === 'inbox' || name === '收件匣' || name === '收件箱') return 'inbox'
  if (lowerName === 'sent items' || lowerName.includes('sent') || name === '寄件備份' || name === '已傳送郵件') return 'sent'
  if (lowerName === 'drafts' || name === '草稿') return 'drafts'
  if (lowerName === 'deleted items' || lowerName.includes('deleted') || name === '刪除的郵件' || name === '垃圾桶') return 'deleted'
  if (lowerName === 'junk email' || lowerName === 'junk e-mail' || name === '垃圾郵件') return 'junk'
  if (lowerName === 'archive' || name === '封存') return 'archive'
  if (lowerName === 'outbox' || name === '寄件匣') return 'outbox'
  return 'normal'
}

export function buildFolderTree(snapshot: FolderSnapshotDto): FolderTreeNode[] {
  return applyFolderBatch([], snapshot.stores, {
    syncId: '',
    sequence: 1,
    reset: true,
    isFinal: true,
    stores: snapshot.stores,
    folders: snapshot.folders,
  })
}

export function applyFolderBatch(current: FolderTreeNode[], stores: OutlookStoreDto[], batch: FolderSyncBatchDto) {
  const roots = batch.reset ? [] : cloneTree(current)
  const byPath = new Map<string, FolderTreeNode>()

  function index(nodes: FolderTreeNode[]) {
    for (const node of nodes) {
      byPath.set(node.folderPath, node)
      index(node.subFolders)
    }
  }

  index(roots)

  for (const folder of batch.folders) {
    const next: FolderTreeNode = {
      ...folder,
      subFolders: byPath.get(folder.folderPath)?.subFolders ?? [],
    }
    byPath.set(next.folderPath, next)

    const siblings = folder.isStoreRoot || !folder.parentFolderPath
      ? roots
      : byPath.get(folder.parentFolderPath)?.subFolders ?? roots
    upsertNode(siblings, next)
  }

  const storeOrder = new Map(stores.map((store, index) => [store.rootFolderPath, index]))
  roots.sort((a, b) => (storeOrder.get(a.folderPath) ?? Number.MAX_SAFE_INTEGER) - (storeOrder.get(b.folderPath) ?? Number.MAX_SAFE_INTEGER))
  return roots
}

function upsertNode(nodes: FolderTreeNode[], next: FolderTreeNode) {
  const index = nodes.findIndex((node) => node.folderPath === next.folderPath)
  if (index < 0) nodes.push(next)
  else nodes[index] = next
}

function cloneTree(nodes: FolderTreeNode[]): FolderTreeNode[] {
  return nodes.map((node) => ({
    ...node,
    subFolders: cloneTree(node.subFolders),
  }))
}
