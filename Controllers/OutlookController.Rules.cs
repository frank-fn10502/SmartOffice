using Microsoft.AspNetCore.Mvc;
using SmartOffice.Hub.Contracts;
using SmartOffice.Hub.Services;

namespace SmartOffice.Hub.Controllers
{
    public partial class OutlookController
    {
        private IActionResult? ValidateRuleRequest(OutlookRuleCommandRequest? req)
        {
            if (req is null)
                return BadRequest(ErrorEnvelope("manage_rule", "missing_rule_request", "rule request is required."));

            NormalizeRuleRequest(req);
            if (req.Operation is not "upsert" and not "delete" and not "set_enabled")
                return BadRequest(ErrorEnvelope("manage_rule", "invalid_rule_operation", "operation must be upsert, delete, or set_enabled."));
            if (req.RuleType is not "receive" and not "send")
                return BadRequest(ErrorEnvelope("manage_rule", "invalid_rule_type", "ruleType must be receive or send."));
            if (string.IsNullOrWhiteSpace(req.RuleName) && string.IsNullOrWhiteSpace(req.OriginalRuleName))
                return BadRequest(ErrorEnvelope("manage_rule", "missing_rule_name", "ruleName or originalRuleName is required."));
            if (req.Conditions.HasAttachment == false)
                return BadRequest(ErrorEnvelope("manage_rule", "unsupported_rule_condition", "Outlook object model only supports the has-attachment rule condition."));
            if (req.Conditions.Importance is not "any" and not "low" and not "normal" and not "high")
                return BadRequest(ErrorEnvelope("manage_rule", "invalid_rule_importance", "importance must be any, low, normal, or high."));
            if (req.Actions.MarkAsTaskInterval is not "today" and not "tomorrow" and not "this_week" and not "next_week" and not "no_date")
                return BadRequest(ErrorEnvelope("manage_rule", "invalid_rule_task_interval", "markAsTaskInterval must be today, tomorrow, this_week, next_week, or no_date."));

            if (req.Operation is "delete" or "set_enabled") return null;
            if (!HasSupportedRuleCondition(req.Conditions))
                return BadRequest(ErrorEnvelope("manage_rule", "missing_rule_condition", "至少需要一個可由 Outlook object model 建立的條件。"));
            if (!HasSupportedRuleAction(req.Actions))
                return BadRequest(ErrorEnvelope("manage_rule", "missing_rule_action", "至少需要一個可由 Outlook object model 建立的動作。"));
            return null;
        }

        private static void NormalizeRuleRequest(OutlookRuleCommandRequest req)
        {
            req.Operation = string.IsNullOrWhiteSpace(req.Operation) ? "upsert" : req.Operation.Trim().ToLowerInvariant();
            req.RuleType = string.IsNullOrWhiteSpace(req.RuleType) ? "receive" : req.RuleType.Trim().ToLowerInvariant();
            req.RuleName = req.RuleName.Trim();
            req.OriginalRuleName = req.OriginalRuleName.Trim();
            req.Conditions ??= new OutlookRuleConditionsRequest();
            req.Actions ??= new OutlookRuleActionsRequest();

            NormalizeRuleList(req.Conditions.SubjectContains);
            NormalizeRuleList(req.Conditions.BodyContains);
            NormalizeRuleList(req.Conditions.BodyOrSubjectContains);
            NormalizeRuleList(req.Conditions.MessageHeaderContains);
            NormalizeRuleList(req.Conditions.SenderAddressContains);
            NormalizeRuleList(req.Conditions.RecipientAddressContains);
            NormalizeRuleList(req.Conditions.Categories);
            NormalizeRuleList(req.Actions.AssignCategories);

            req.Conditions.Importance = string.IsNullOrWhiteSpace(req.Conditions.Importance)
                ? "any"
                : req.Conditions.Importance.Trim().ToLowerInvariant();
            req.Actions.MarkAsTaskInterval = string.IsNullOrWhiteSpace(req.Actions.MarkAsTaskInterval)
                ? "today"
                : req.Actions.MarkAsTaskInterval.Trim().ToLowerInvariant();
            req.Actions.MoveToFolderPath = OutlookFolderPathMapper.ToAddinPath(req.Actions.MoveToFolderPath.Trim());
            req.Actions.CopyToFolderPath = OutlookFolderPathMapper.ToAddinPath(req.Actions.CopyToFolderPath.Trim());
        }

        private static bool HasSupportedRuleCondition(OutlookRuleConditionsRequest conditions)
        {
            return conditions.SubjectContains.Count > 0
                || conditions.BodyContains.Count > 0
                || conditions.BodyOrSubjectContains.Count > 0
                || conditions.MessageHeaderContains.Count > 0
                || conditions.SenderAddressContains.Count > 0
                || conditions.RecipientAddressContains.Count > 0
                || conditions.Categories.Count > 0
                || conditions.HasAttachment is not null
                || conditions.Importance != "any"
                || conditions.ToMe
                || conditions.ToOrCcMe
                || conditions.OnlyToMe
                || conditions.MeetingInviteOrUpdate;
        }

        private static bool HasSupportedRuleAction(OutlookRuleActionsRequest actions)
        {
            return !string.IsNullOrWhiteSpace(actions.MoveToFolderPath)
                || !string.IsNullOrWhiteSpace(actions.CopyToFolderPath)
                || actions.AssignCategories.Count > 0
                || actions.ClearCategories
                || actions.MarkAsTask
                || actions.Delete
                || actions.DesktopAlert
                || actions.StopProcessingMoreRules;
        }

        private static void NormalizeRuleList(List<string> values)
        {
            var normalized = values
                .Select(value => value.Trim())
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            values.Clear();
            values.AddRange(normalized);
        }
    }
}
