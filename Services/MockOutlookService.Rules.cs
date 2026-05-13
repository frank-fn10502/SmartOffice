namespace SmartOffice.Hub.Services
{
    public partial class MockOutlookService
    {
        private static List<string> BuildRuleConditionSummaries(OutlookRuleConditionsRequest conditions)
        {
            var result = new List<string>();
            if (conditions.SubjectContains.Count > 0) result.Add($"Subject: Text={string.Join(", ", conditions.SubjectContains)}");
            if (conditions.BodyContains.Count > 0) result.Add($"Body: Text={string.Join(", ", conditions.BodyContains)}");
            if (conditions.BodyOrSubjectContains.Count > 0) result.Add($"BodyOrSubject: Text={string.Join(", ", conditions.BodyOrSubjectContains)}");
            if (conditions.MessageHeaderContains.Count > 0) result.Add($"MessageHeader: Text={string.Join(", ", conditions.MessageHeaderContains)}");
            if (conditions.SenderAddressContains.Count > 0) result.Add($"SenderAddress: Address={string.Join(", ", conditions.SenderAddressContains)}");
            if (conditions.RecipientAddressContains.Count > 0) result.Add($"RecipientAddress: Address={string.Join(", ", conditions.RecipientAddressContains)}");
            if (conditions.Categories.Count > 0) result.Add($"Category: Categories={string.Join(", ", conditions.Categories)}");
            if (conditions.HasAttachment == true) result.Add("HasAttachment: (enabled)");
            if (conditions.Importance is "low" or "normal" or "high") result.Add($"Importance: Importance={conditions.Importance}");
            if (conditions.ToMe) result.Add("ToMe: (enabled)");
            if (conditions.ToOrCcMe) result.Add("ToOrCc: (enabled)");
            if (conditions.OnlyToMe) result.Add("OnlyToMe: (enabled)");
            if (conditions.MeetingInviteOrUpdate) result.Add("MeetingInviteOrUpdate: (enabled)");
            return result;
        }

        private static List<string> BuildRuleActionSummaries(OutlookRuleActionsRequest actions)
        {
            var result = new List<string>();
            if (!string.IsNullOrWhiteSpace(actions.MoveToFolderPath)) result.Add($"MoveToFolder: FolderPath={actions.MoveToFolderPath}");
            if (!string.IsNullOrWhiteSpace(actions.CopyToFolderPath)) result.Add($"CopyToFolder: FolderPath={actions.CopyToFolderPath}");
            if (actions.AssignCategories.Count > 0) result.Add($"AssignToCategory: Categories={string.Join(", ", actions.AssignCategories)}");
            if (actions.ClearCategories) result.Add("ClearCategories: (enabled)");
            if (actions.MarkAsTask) result.Add($"MarkAsTask: MarkInterval={actions.MarkAsTaskInterval}");
            if (actions.Delete) result.Add("Delete: (enabled)");
            if (actions.DesktopAlert) result.Add("DesktopAlert: (enabled)");
            if (actions.StopProcessingMoreRules) result.Add("Stop: (enabled)");
            return result;
        }
    }
}
