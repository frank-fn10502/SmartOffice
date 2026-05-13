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
  senderAddressContains: string
  categories: string[]
  hasAttachment: 'any' | 'yes'
  moveToFolderPath: string
  assignCategories: string[]
  markAsTask: boolean
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
    senderAddressContains: '',
    categories: [],
    hasAttachment: 'any',
    moveToFolderPath: '',
    assignCategories: [],
    markAsTask: false,
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
    senderAddressContains: firstRuleSummaryList(rule.conditions, 'SenderAddress:', 'Address').join(', '),
    categories: firstRuleSummaryList(rule.conditions, 'Category:', 'Categories'),
    hasAttachment: rule.conditions.some((condition) => condition.toLowerCase().startsWith('hasattachment:')) ? 'yes' : 'any',
    moveToFolderPath: firstRuleSummaryValue(rule.actions, 'MoveToFolder:', 'FolderPath'),
    assignCategories: firstRuleSummaryList(rule.actions, 'AssignToCategory:', 'Categories'),
    markAsTask: rule.actions.some((action) => action.toLowerCase().includes('task')),
    stopProcessingMoreRules: rule.actions.some((action) => action.toLowerCase().includes('stop')),
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
    senderAddressContains: '',
    categories: [],
    hasAttachment: 'any',
    moveToFolderPath: '',
    assignCategories: [],
    markAsTask: false,
    stopProcessingMoreRules: false,
  }
}
