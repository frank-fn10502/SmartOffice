import type { FolderDto } from '../models/outlook'

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

export function visibleChildren(folder: FolderDto) {
  return (folder.subFolders ?? []).filter((child) => !isHiddenFolder(child.name))
}

export function visibleRootFolders(folders: FolderDto[]) {
  return folders.flatMap((root) => {
    if (root.subFolders?.length) return root.subFolders.filter((folder) => !isHiddenFolder(folder.name))
    return isHiddenFolder(root.name) ? [] : [root]
  })
}

export function collectFolderOptions(items: FolderDto[], level = 0): Array<FolderDto & { label: string }> {
  return items.flatMap((folder) => {
    const indent = level === 0 ? '' : `${'　'.repeat(level)}`
    return [
      { ...folder, label: `${indent}${folder.name}` },
      ...collectFolderOptions(visibleChildren(folder), level + 1),
    ]
  })
}

export function findFolderByPath(items: FolderDto[], path: string): FolderDto | null {
  for (const folder of items) {
    if (folder.folderPath === path) return folder
    const child = findFolderByPath(visibleChildren(folder), path)
    if (child) return child
  }
  return null
}

export function folderType(name: string) {
  const lowerName = name.toLowerCase()
  if (lowerName === 'inbox') return 'inbox'
  if (lowerName === 'sent items' || lowerName.includes('sent')) return 'sent'
  if (lowerName === 'drafts') return 'drafts'
  if (lowerName === 'deleted items' || lowerName.includes('deleted')) return 'deleted'
  if (lowerName === 'junk email' || lowerName === 'junk e-mail') return 'junk'
  if (lowerName === 'archive') return 'archive'
  if (lowerName === 'outbox') return 'outbox'
  return 'normal'
}
