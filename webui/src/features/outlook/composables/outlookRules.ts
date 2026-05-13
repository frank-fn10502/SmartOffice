import type { OutlookRuleCommandRequest, OutlookRuleDto } from '../models/outlook'

export type RuleDraft = {
  storeId: string
  ruleName: string
  originalRuleName: string
  originalExecutionOrder?: number
  ruleType: 'receive' | 'send'
  enabled: boolean
  subjectContains: string
  bodyContains: string
  bodyOrSubjectContains: string
  messageHeaderContains: string
  senderAddressContains: string
  recipientAddressContains: string
  categories: string[]
  hasAttachment: 'any' | 'yes'
  importance: 'any' | 'low' | 'normal' | 'high'
  toMe: boolean
  toOrCcMe: boolean
  onlyToMe: boolean
  meetingInviteOrUpdate: boolean
  moveToFolderPath: string
  copyToFolderPath: string
  assignCategories: string[]
  clearCategories: boolean
  markAsTask: boolean
  markAsTaskInterval: 'today' | 'tomorrow' | 'this_week' | 'next_week' | 'no_date'
  delete: boolean
  desktopAlert: boolean
  stopProcessingMoreRules: boolean
}

export function createEmptyRuleDraft(): RuleDraft {
  return {
    ruleName: '',
    storeId: '',
    originalRuleName: '',
    originalExecutionOrder: undefined,
    ruleType: 'receive',
    enabled: true,
    subjectContains: '',
    bodyContains: '',
    bodyOrSubjectContains: '',
    messageHeaderContains: '',
    senderAddressContains: '',
    recipientAddressContains: '',
    categories: [],
    hasAttachment: 'any',
    importance: 'any',
    toMe: false,
    toOrCcMe: false,
    onlyToMe: false,
    meetingInviteOrUpdate: false,
    moveToFolderPath: '',
    copyToFolderPath: '',
    assignCategories: [],
    clearCategories: false,
    markAsTask: false,
    markAsTaskInterval: 'today',
    delete: false,
    desktopAlert: false,
    stopProcessingMoreRules: true,
  }
}

export function splitRuleInput(value: string) {
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

function hasRuleSummary(summaries: string[], prefix: string) {
  return summaries.some((item) => item.toLowerCase().startsWith(prefix.toLowerCase()))
}

function firstRuleImportance(summaries: string[]) {
  const value = firstRuleSummaryValue(summaries, 'Importance:', 'Importance').toLowerCase()
  return value === 'low' || value === 'normal' || value === 'high' ? value : 'any'
}

function firstTaskInterval(summaries: string[]) {
  const value = firstRuleSummaryValue(summaries, 'MarkAsTask:', 'MarkInterval').toLowerCase()
  if (value === 'tomorrow' || value === 'this_week' || value === 'next_week' || value === 'no_date') return value
  return 'today'
}

export function buildRuleDraftFromRule(rule: OutlookRuleDto): RuleDraft {
  return {
    ruleName: rule.name,
    storeId: rule.storeId,
    originalRuleName: rule.name,
    originalExecutionOrder: rule.executionOrder,
    ruleType: rule.ruleType?.toLowerCase() === 'send' ? 'send' : 'receive',
    enabled: rule.enabled,
    subjectContains: firstRuleSummaryList(rule.conditions, 'Subject:', 'Text').join(', '),
    bodyContains: firstRuleSummaryList(rule.conditions, 'Body:', 'Text').join(', '),
    bodyOrSubjectContains: firstRuleSummaryList(rule.conditions, 'BodyOrSubject:', 'Text').join(', '),
    messageHeaderContains: firstRuleSummaryList(rule.conditions, 'MessageHeader:', 'Text').join(', '),
    senderAddressContains: firstRuleSummaryList(rule.conditions, 'SenderAddress:', 'Address').join(', '),
    recipientAddressContains: firstRuleSummaryList(rule.conditions, 'RecipientAddress:', 'Address').join(', '),
    categories: firstRuleSummaryList(rule.conditions, 'Category:', 'Categories'),
    hasAttachment: hasRuleSummary(rule.conditions, 'HasAttachment:') ? 'yes' : 'any',
    importance: firstRuleImportance(rule.conditions),
    toMe: hasRuleSummary(rule.conditions, 'ToMe:'),
    toOrCcMe: hasRuleSummary(rule.conditions, 'ToOrCc:'),
    onlyToMe: hasRuleSummary(rule.conditions, 'OnlyToMe:'),
    meetingInviteOrUpdate: hasRuleSummary(rule.conditions, 'MeetingInviteOrUpdate:'),
    moveToFolderPath: firstRuleSummaryValue(rule.actions, 'MoveToFolder:', 'FolderPath'),
    copyToFolderPath: firstRuleSummaryValue(rule.actions, 'CopyToFolder:', 'FolderPath'),
    assignCategories: firstRuleSummaryList(rule.actions, 'AssignToCategory:', 'Categories'),
    clearCategories: hasRuleSummary(rule.actions, 'ClearCategories:'),
    markAsTask: hasRuleSummary(rule.actions, 'MarkAsTask:'),
    markAsTaskInterval: firstTaskInterval(rule.actions),
    delete: hasRuleSummary(rule.actions, 'Delete:'),
    desktopAlert: hasRuleSummary(rule.actions, 'DesktopAlert:'),
    stopProcessingMoreRules: hasRuleSummary(rule.actions, 'Stop:'),
  }
}

export function buildRulePayload(operation: 'upsert' | 'delete' | 'set_enabled', draft: RuleDraft): OutlookRuleCommandRequest {
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
      bodyOrSubjectContains: splitRuleInput(draft.bodyOrSubjectContains),
      messageHeaderContains: splitRuleInput(draft.messageHeaderContains),
      senderAddressContains: splitRuleInput(draft.senderAddressContains),
      recipientAddressContains: splitRuleInput(draft.recipientAddressContains),
      categories: draft.categories,
      hasAttachment,
      importance: draft.importance,
      toMe: draft.toMe,
      toOrCcMe: draft.toOrCcMe,
      onlyToMe: draft.onlyToMe,
      meetingInviteOrUpdate: draft.meetingInviteOrUpdate,
    },
    actions: {
      moveToFolderPath: draft.moveToFolderPath,
      copyToFolderPath: draft.copyToFolderPath,
      assignCategories: draft.assignCategories,
      clearCategories: draft.clearCategories,
      markAsTask: draft.markAsTask,
      markAsTaskInterval: draft.markAsTaskInterval,
      delete: draft.delete,
      desktopAlert: draft.desktopAlert,
      stopProcessingMoreRules: draft.stopProcessingMoreRules,
    },
  }
}

export function buildRuleOperationDraft(rule: OutlookRuleDto, enabled = rule.enabled): RuleDraft {
  return {
    ruleName: rule.name,
    storeId: rule.storeId,
    originalRuleName: rule.name,
    originalExecutionOrder: rule.executionOrder,
    ruleType: rule.ruleType?.toLowerCase() === 'send' ? 'send' : 'receive',
    enabled,
    subjectContains: '',
    bodyContains: '',
    bodyOrSubjectContains: '',
    messageHeaderContains: '',
    senderAddressContains: '',
    recipientAddressContains: '',
    categories: [],
    hasAttachment: 'any',
    importance: 'any',
    toMe: false,
    toOrCcMe: false,
    onlyToMe: false,
    meetingInviteOrUpdate: false,
    moveToFolderPath: '',
    copyToFolderPath: '',
    assignCategories: [],
    clearCategories: false,
    markAsTask: false,
    markAsTaskInterval: 'today',
    delete: false,
    desktopAlert: false,
    stopProcessingMoreRules: false,
  }
}
