using System;
using System.Collections.Generic;

namespace SmartOffice.Hub.Contracts
{
    public class FolderSyncBeginDto
    {
        public string SyncId { get; set; } = Guid.NewGuid().ToString();
        public bool Reset { get; set; } = true;
        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
    }

    public class FolderSyncBatchDto
    {
        public string SyncId { get; set; } = string.Empty;
        public int Sequence { get; set; }
        public bool Reset { get; set; }
        public bool IsFinal { get; set; }
        public List<OutlookStoreDto> Stores { get; set; } = new List<OutlookStoreDto>();
        public List<FolderDto> Folders { get; set; } = new List<FolderDto>();
    }

    public class FolderSyncCompleteDto
    {
        public string SyncId { get; set; } = string.Empty;
        public int TotalCount { get; set; }
        public bool Success { get; set; } = true;
        public string Message { get; set; } = string.Empty;
        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
    }

    public class FolderDiscoveryRequest
    {
        public string SyncId { get; set; } = Guid.NewGuid().ToString();
        public string StoreId { get; set; } = string.Empty;
        public string ParentEntryId { get; set; } = string.Empty;
        public string ParentFolderPath { get; set; } = string.Empty;
        public int MaxDepth { get; set; } = 1;
        public int MaxChildren { get; set; } = 50;
        public bool Reset { get; set; }
    }

    public class FetchMailsRequest
    {
        public string FolderPath { get; set; } = string.Empty;
        public string Range { get; set; } = string.Empty;
        public DateTime? ReceivedFrom { get; set; }
        public DateTime? ReceivedTo { get; set; }
        public int MaxCount { get; set; } = 30;
    }

    public class SearchMailsRequest
    {
        public string SearchId { get; set; } = Guid.NewGuid().ToString();
        public string StoreId { get; set; } = string.Empty;
        public List<string> ScopeFolderPaths { get; set; } = new List<string>();
        public bool AllowGlobalScope { get; set; }
        public bool IncludeSubFolders { get; set; } = true;
        public string Keyword { get; set; } = string.Empty;
        public List<string> TextFields { get; set; } = new List<string> { "subject" };
        public List<string> CategoryNames { get; set; } = new List<string>();
        public bool? HasAttachments { get; set; }
        public string FlagState { get; set; } = "any";
        public string ReadState { get; set; } = "any";
        public DateTime? ReceivedFrom { get; set; }
        public DateTime? ReceivedTo { get; set; }
    }

    public class MailSearchSliceRequest
    {
        public string SearchId { get; set; } = string.Empty;
        public string CommandId { get; set; } = string.Empty;
        public string ParentCommandId { get; set; } = string.Empty;
        public string StoreId { get; set; } = string.Empty;
        public string FolderEntryId { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
        public string ExecutionMode { get; set; } = "items_filter";
        public string Keyword { get; set; } = string.Empty;
        public List<string> TextFields { get; set; } = new List<string> { "subject" };
        public List<string> CategoryNames { get; set; } = new List<string>();
        public bool? HasAttachments { get; set; }
        public string FlagState { get; set; } = "any";
        public string ReadState { get; set; } = "any";
        public DateTime? ReceivedFrom { get; set; }
        public DateTime? ReceivedTo { get; set; }
        public int SliceIndex { get; set; }
        public int SliceCount { get; set; }
        public int ResultBatchSize { get; set; } = 5;
        public bool ResetSearchResults { get; set; } = true;
        public bool CompleteSearchOnSlice { get; set; } = true;
    }

    public class FolderMailsSliceRequest
    {
        public string FolderMailsId { get; set; } = string.Empty;
        public string CommandId { get; set; } = string.Empty;
        public string ParentCommandId { get; set; } = string.Empty;
        public string StoreId { get; set; } = string.Empty;
        public string FolderEntryId { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
        public DateTime? ReceivedFrom { get; set; }
        public DateTime? ReceivedTo { get; set; }
        public int MaxCount { get; set; } = 30;
        public int SliceIndex { get; set; }
        public int SliceCount { get; set; }
        public int ResultBatchSize { get; set; } = 5;
        public bool ResetResults { get; set; } = true;
        public bool CompleteOnSlice { get; set; } = true;
    }

    public class MailSearchSliceResultDto
    {
        public string SearchId { get; set; } = string.Empty;
        public string CommandId { get; set; } = string.Empty;
        public string ParentCommandId { get; set; } = string.Empty;
        public int Sequence { get; set; }
        public int SliceIndex { get; set; }
        public int SliceCount { get; set; }
        public bool Reset { get; set; }
        public bool IsFinal { get; set; }
        public bool IsSliceComplete { get; set; } = true;
        public List<MailItemDto> Mails { get; set; } = new List<MailItemDto>();
        public string Message { get; set; } = string.Empty;
    }

    public class MailSearchCompleteDto
    {
        public string SearchId { get; set; } = string.Empty;
        public string CommandId { get; set; } = string.Empty;
        public string ParentCommandId { get; set; } = string.Empty;
        public int TotalCount { get; set; }
        public bool Success { get; set; } = true;
        public string Message { get; set; } = string.Empty;
        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
    }

    public class FolderMailsSliceResultDto
    {
        public string FolderMailsId { get; set; } = string.Empty;
        public string CommandId { get; set; } = string.Empty;
        public string ParentCommandId { get; set; } = string.Empty;
        public int Sequence { get; set; }
        public int SliceIndex { get; set; }
        public int SliceCount { get; set; }
        public bool Reset { get; set; }
        public bool IsFinal { get; set; }
        public bool IsSliceComplete { get; set; } = true;
        public List<MailItemDto> Mails { get; set; } = new List<MailItemDto>();
        public string Message { get; set; } = string.Empty;
    }

    public class FolderMailsCompleteDto
    {
        public string FolderMailsId { get; set; } = string.Empty;
        public string CommandId { get; set; } = string.Empty;
        public string ParentCommandId { get; set; } = string.Empty;
        public int TotalCount { get; set; }
        public bool Success { get; set; } = true;
        public string Message { get; set; } = string.Empty;
        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
    }

    public class FetchCalendarRequest
    {
        public int DaysForward { get; set; } = 31;
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }
    }

    public class FetchMailBodyRequest
    {
        public string MailId { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
    }

    public class FetchMailAttachmentsRequest
    {
        public string MailId { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
    }

    public class FetchMailConversationRequest
    {
        public string MailId { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
        public int MaxCount { get; set; } = 100;
        public bool IncludeBody { get; set; } = true;
    }

    public class ExportMailAttachmentRequest
    {
        public string MailId { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
        public string AttachmentId { get; set; } = string.Empty;
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string ExportRootPath { get; set; } = string.Empty;
    }

    public class MailPropertiesCommandRequest
    {
        public string MailId { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
        public bool? IsRead { get; set; }
        public string FlagInterval { get; set; } = "none";
        public string FlagRequest { get; set; } = string.Empty;
        public DateTime? TaskStartDate { get; set; }
        public DateTime? TaskDueDate { get; set; }
        public DateTime? TaskCompletedDate { get; set; }
        public List<string> Categories { get; set; } = new List<string>();
        public List<OutlookCategoryDto> NewCategories { get; set; } = new List<OutlookCategoryDto>();
    }

    public class CategoryCommandRequest
    {
        public string Name { get; set; } = string.Empty;
        public string Color { get; set; } = string.Empty;
        public int ColorValue { get; set; }
        public string ShortcutKey { get; set; } = string.Empty;
    }

    public class OutlookRuleConditionsRequest
    {
        public List<string> SubjectContains { get; set; } = new List<string>();
        public List<string> BodyContains { get; set; } = new List<string>();
        public List<string> BodyOrSubjectContains { get; set; } = new List<string>();
        public List<string> MessageHeaderContains { get; set; } = new List<string>();
        public List<string> SenderAddressContains { get; set; } = new List<string>();
        public List<string> RecipientAddressContains { get; set; } = new List<string>();
        public List<string> Categories { get; set; } = new List<string>();
        public bool? HasAttachment { get; set; }
        public string Importance { get; set; } = "any";
        public bool ToMe { get; set; }
        public bool ToOrCcMe { get; set; }
        public bool OnlyToMe { get; set; }
        public bool MeetingInviteOrUpdate { get; set; }
    }

    public class OutlookRuleActionsRequest
    {
        public string MoveToFolderPath { get; set; } = string.Empty;
        public string CopyToFolderPath { get; set; } = string.Empty;
        public List<string> AssignCategories { get; set; } = new List<string>();
        public bool ClearCategories { get; set; }
        public bool MarkAsTask { get; set; }
        public string MarkAsTaskInterval { get; set; } = "today";
        public bool Delete { get; set; }
        public bool DesktopAlert { get; set; }
        public bool StopProcessingMoreRules { get; set; } = true;
    }

    public class OutlookRuleCommandRequest
    {
        public string Operation { get; set; } = "upsert";
        public string StoreId { get; set; } = string.Empty;
        public string RuleName { get; set; } = string.Empty;
        public string OriginalRuleName { get; set; } = string.Empty;
        public int? OriginalExecutionOrder { get; set; }
        public string RuleType { get; set; } = "receive";
        public bool Enabled { get; set; } = true;
        public int? ExecutionOrder { get; set; }
        public OutlookRuleConditionsRequest Conditions { get; set; } = new OutlookRuleConditionsRequest();
        public OutlookRuleActionsRequest Actions { get; set; } = new OutlookRuleActionsRequest();
    }

    public class CreateFolderRequest
    {
        public string ParentFolderPath { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
    }

    public class DeleteFolderRequest
    {
        public string FolderPath { get; set; } = string.Empty;
    }

    public class MoveMailRequest
    {
        public string MailId { get; set; } = string.Empty;
        public string SourceFolderPath { get; set; } = string.Empty;
        public string DestinationFolderPath { get; set; } = string.Empty;
    }

    public class MoveMailsRequest
    {
        public List<string> MailIds { get; set; } = new List<string>();
        public string SourceFolderPath { get; set; } = string.Empty;
        public List<string> SourceFolderPaths { get; set; } = new List<string>();
        public string DestinationFolderPath { get; set; } = string.Empty;
        public bool ContinueOnError { get; set; } = true;
    }

    public class DeleteMailRequest
    {
        public string MailId { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
    }

    public class FindFolderRequest
    {
        public string Name { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
        public string FolderType { get; set; } = string.Empty;
        public string StoreId { get; set; } = string.Empty;
        public bool IncludeHidden { get; set; }
        public int MaxResults { get; set; } = 20;
    }

    public class PendingCommand
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Type { get; set; } = string.Empty;
        public FolderDiscoveryRequest FolderDiscoveryRequest { get; set; }
        public FindFolderRequest FindFolderRequest { get; set; }
        public FetchMailsRequest MailsRequest { get; set; }
        public SearchMailsRequest SearchMailsRequest { get; set; }
        public MailSearchSliceRequest MailSearchSliceRequest { get; set; }
        public FolderMailsSliceRequest FolderMailsSliceRequest { get; set; }
        public FetchMailBodyRequest MailBodyRequest { get; set; }
        public FetchMailAttachmentsRequest MailAttachmentsRequest { get; set; }
        public FetchMailConversationRequest MailConversationRequest { get; set; }
        public ExportMailAttachmentRequest ExportMailAttachmentRequest { get; set; }
        public FetchCalendarRequest CalendarRequest { get; set; }
        public MailPropertiesCommandRequest MailPropertiesRequest { get; set; }
        public CategoryCommandRequest CategoryRequest { get; set; }
        public OutlookRuleCommandRequest RuleRequest { get; set; }
        public CreateFolderRequest CreateFolderRequest { get; set; }
        public DeleteFolderRequest DeleteFolderRequest { get; set; }
        public MoveMailRequest MoveMailRequest { get; set; }
        public MoveMailsRequest MoveMailsRequest { get; set; }
        public DeleteMailRequest DeleteMailRequest { get; set; }
        public AddressBookSyncRequest AddressBookRequest { get; set; }
    }

    public class OutlookAddinClientInfo
    {
        public string ClientName { get; set; } = string.Empty;
        public string Workstation { get; set; } = string.Empty;
        public string Version { get; set; } = string.Empty;
    }

    public class OutlookCommandResult
    {
        public string CommandId { get; set; } = string.Empty;
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public string Payload { get; set; } = string.Empty;
        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
    }

    public class OutlookCommand
    {
        public string Id { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public OutlookCommandFolderDiscoveryRequest FolderDiscoveryRequest { get; set; }
        public OutlookCommandMailsRequest MailsRequest { get; set; }
        public OutlookCommandMailSearchSliceRequest MailSearchSliceRequest { get; set; }
        public OutlookCommandMailBodyRequest MailBodyRequest { get; set; }
        public OutlookCommandMailAttachmentsRequest MailAttachmentsRequest { get; set; }
        public OutlookCommandMailConversationRequest MailConversationRequest { get; set; }
        public OutlookCommandExportMailAttachmentRequest ExportMailAttachmentRequest { get; set; }
        public OutlookCommandCalendarRequest CalendarRequest { get; set; }
        public OutlookCommandMailPropertiesRequest MailPropertiesRequest { get; set; }
        public OutlookCommandCategoryRequest CategoryRequest { get; set; }
        public OutlookCommandCreateFolderRequest CreateFolderRequest { get; set; }
        public OutlookCommandDeleteFolderRequest DeleteFolderRequest { get; set; }
        public OutlookCommandMoveMailRequest MoveMailRequest { get; set; }
        public OutlookCommandMoveMailsRequest MoveMailsRequest { get; set; }
        public OutlookCommandDeleteMailRequest DeleteMailRequest { get; set; }
        public OutlookCommandFolderMailsSliceRequest FolderMailsSliceRequest { get; set; }
        public OutlookCommandRuleRequest RuleRequest { get; set; }
        public OutlookCommandAddressBookRequest AddressBookRequest { get; set; }
    }

    public class OutlookCommandMailsRequest : FetchMailsRequest { }
    public class OutlookCommandMailSearchSliceRequest : MailSearchSliceRequest { }
    public class OutlookCommandFolderDiscoveryRequest : FolderDiscoveryRequest { }
    public class OutlookCommandMoveMailsRequest : MoveMailsRequest { }
    public class OutlookCommandFolderMailsSliceRequest : FolderMailsSliceRequest { }
    public class OutlookCommandMailBodyRequest : FetchMailBodyRequest { }
    public class OutlookCommandMailAttachmentsRequest : FetchMailAttachmentsRequest { }
    public class OutlookCommandMailConversationRequest : FetchMailConversationRequest { }
    public class OutlookCommandExportMailAttachmentRequest : ExportMailAttachmentRequest { }
    public class OutlookCommandMailPropertiesRequest : MailPropertiesCommandRequest { }
    public class OutlookCommandCategoryRequest : CategoryCommandRequest { }
    public class OutlookCommandCreateFolderRequest : CreateFolderRequest { }
    public class OutlookCommandDeleteFolderRequest : DeleteFolderRequest { }
    public class OutlookCommandMoveMailRequest : MoveMailRequest { }
    public class OutlookCommandDeleteMailRequest : DeleteMailRequest { }
    public class OutlookCommandAddressBookRequest : AddressBookSyncRequest { }
    public class OutlookCommandRuleConditions : OutlookRuleConditionsRequest { }
    public class OutlookCommandRuleActions : OutlookRuleActionsRequest { }

    public class OutlookCommandCalendarRequest
    {
        public int DaysForward { get; set; } = 31;
        public string StartDate { get; set; } = string.Empty;
        public string EndDate { get; set; } = string.Empty;
    }

    public class OutlookCommandNewCategory : OutlookCategoryDto { }

    public class OutlookCommandRuleRequest
    {
        public string Operation { get; set; } = "upsert";
        public string StoreId { get; set; } = string.Empty;
        public string RuleName { get; set; } = string.Empty;
        public string OriginalRuleName { get; set; } = string.Empty;
        public int? OriginalExecutionOrder { get; set; }
        public string RuleType { get; set; } = "receive";
        public bool Enabled { get; set; } = true;
        public int? ExecutionOrder { get; set; }
        public OutlookCommandRuleConditions Conditions { get; set; } = new OutlookCommandRuleConditions();
        public OutlookCommandRuleActions Actions { get; set; } = new OutlookCommandRuleActions();
    }
}
