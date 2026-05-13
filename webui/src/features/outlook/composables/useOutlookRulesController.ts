import type { Ref } from 'vue'
import { ElMessage } from 'element-plus'
import { outlookApi } from '../api/outlook'
import type { OutlookRuleDto } from '../models/outlook'
import { categoryColorValue } from '../utils/categoryColors'
import {
  buildRuleDraftFromRule,
  buildRuleOperationDraft,
  buildRulePayload,
  createEmptyRuleDraft,
  splitRuleInput,
  type RuleDraft,
} from './outlookRules'

type RulesControllerOptions = {
  rules: Ref<OutlookRuleDto[]>
  selectedRuleIndex: Ref<number | null>
  ruleDraft: Ref<RuleDraft>
  loadingRules: Ref<boolean>
  outlookBusy: Ref<boolean>
  runMailOperation: (action: () => Promise<unknown>, afterSuccess?: () => Promise<void>) => Promise<boolean>
  loadCachedRules: () => Promise<void>
  loadCachedCategories: () => Promise<void>
}

export function useOutlookRulesController(options: RulesControllerOptions) {
  const {
    loadCachedCategories,
    loadCachedRules,
    loadingRules,
    outlookBusy,
    ruleDraft,
    rules,
    runMailOperation,
    selectedRuleIndex,
  } = options

  async function requestRules() {
    if (loadingRules.value) return
    if (outlookBusy.value) return
    loadingRules.value = true
    try {
      const response = await outlookApi.requestRules()
      await runMailOperation(
        async () => response,
        async () => loadCachedRules(),
      )
      loadingRules.value = false
    } catch {
      loadingRules.value = false
    }
  }

  function resetRuleDraft(rule: OutlookRuleDto | null = null) {
    if (!rule) {
      ruleDraft.value = createEmptyRuleDraft()
      selectedRuleIndex.value = null
      return
    }

    ruleDraft.value = buildRuleDraftFromRule(rule)
  }

  function editRule(index: number) {
    const rule = rules.value[index]
    if (!rule) return
    selectedRuleIndex.value = index
    resetRuleDraft(rule)
  }

  async function saveRule() {
    if (outlookBusy.value || !ruleDraft.value.ruleName.trim()) return false
    const hasCondition = splitRuleInput(ruleDraft.value.subjectContains).length > 0
      || splitRuleInput(ruleDraft.value.bodyContains).length > 0
      || splitRuleInput(ruleDraft.value.bodyOrSubjectContains).length > 0
      || splitRuleInput(ruleDraft.value.messageHeaderContains).length > 0
      || splitRuleInput(ruleDraft.value.senderAddressContains).length > 0
      || splitRuleInput(ruleDraft.value.recipientAddressContains).length > 0
      || ruleDraft.value.categories.length > 0
      || ruleDraft.value.hasAttachment === 'yes'
      || ruleDraft.value.importance !== 'any'
      || ruleDraft.value.toMe
      || ruleDraft.value.toOrCcMe
      || ruleDraft.value.onlyToMe
      || ruleDraft.value.meetingInviteOrUpdate
    const hasAction = Boolean(ruleDraft.value.moveToFolderPath)
      || Boolean(ruleDraft.value.copyToFolderPath)
      || ruleDraft.value.assignCategories.length > 0
      || ruleDraft.value.clearCategories
      || ruleDraft.value.markAsTask
      || ruleDraft.value.delete
      || ruleDraft.value.desktopAlert
      || ruleDraft.value.stopProcessingMoreRules
    if (!hasCondition || !hasAction) {
      ElMessage.warning('請至少設定一個條件與一個動作。')
      return false
    }
    await runMailOperation(
      () => outlookApi.requestManageRule(buildRulePayload('upsert', ruleDraft.value)),
      async () => {
        await Promise.allSettled([loadCachedRules(), loadCachedCategories()])
        resetRuleDraft()
      },
    )
    return true
  }

  async function deleteRule(rule = selectedRuleIndex.value === null ? null : rules.value[selectedRuleIndex.value]) {
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
    if (!rule || outlookBusy.value) return
    const payload = buildRulePayload('set_enabled', buildRuleOperationDraft(rule, enabled))
    await runMailOperation(
      () => outlookApi.requestManageRule(payload),
      async () => loadCachedRules(),
    )
  }

  async function upsertCategory(name: string, color: string, shortcutKey = '') {
    if (!name.trim()) return false
    await runMailOperation(
      () => outlookApi.requestUpsertCategory({
        name: name.trim(),
        color,
        colorValue: categoryColorValue(color || 'olCategoryColorNone'),
        shortcutKey,
      }),
      async () => loadCachedCategories(),
    )
    return true
  }

  return {
    deleteRule,
    editRule,
    requestRules,
    resetRuleDraft,
    saveRule,
    toggleRuleEnabled,
    upsertCategory,
  }
}
