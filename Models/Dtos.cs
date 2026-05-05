namespace SmartOffice.Hub.Models
{
    public class MailItemDto
    {
        public string Id { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public string SenderName { get; set; } = string.Empty;
        public string SenderEmail { get; set; } = string.Empty;
        public DateTime ReceivedTime { get; set; }
        public string Body { get; set; } = string.Empty;
        public string BodyHtml { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
        public string Categories { get; set; } = string.Empty;
        public bool IsRead { get; set; }
        public bool IsMarkedAsTask { get; set; }
        public int AttachmentCount { get; set; }
        public string AttachmentNames { get; set; } = string.Empty;
        public string FlagRequest { get; set; } = string.Empty;
        public string FlagInterval { get; set; } = "none";
        public DateTime? TaskStartDate { get; set; }
        public DateTime? TaskDueDate { get; set; }
        public DateTime? TaskCompletedDate { get; set; }
        public string Importance { get; set; } = "normal";
        public string Sensitivity { get; set; } = "normal";
    }

    public class ChatMessageDto
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Source { get; set; } = "outlook"; // 目前預期值："outlook" 或 "web"。
        public string Text { get; set; } = string.Empty;
        public DateTime Timestamp { get; set; } = DateTime.Now;
    }

    public class FolderDto
    {
        public string Name { get; set; } = string.Empty;
        public string EntryId { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
        public string ParentEntryId { get; set; } = string.Empty;
        public string ParentFolderPath { get; set; } = string.Empty;
        public int ItemCount { get; set; }
        public string StoreId { get; set; } = string.Empty;
        public bool IsStoreRoot { get; set; }
        public bool HasChildren { get; set; }
        public bool ChildrenLoaded { get; set; }
        public string DiscoveryState { get; set; } = "partial"; // partial、loaded、failed。
    }

    public class OutlookStoreDto
    {
        public string StoreId { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string StoreKind { get; set; } = string.Empty; // ost、pst、exchange、other。
        public string StoreFilePath { get; set; } = string.Empty;
        public string RootFolderPath { get; set; } = string.Empty;
    }

    public class FolderSnapshotDto
    {
        public List<OutlookStoreDto> Stores { get; set; } = new();
        public List<FolderDto> Folders { get; set; } = new();
    }

    public class FolderSyncBeginDto
    {
        public string SyncId { get; set; } = Guid.NewGuid().ToString();
        public bool Reset { get; set; } = true;
        public DateTime Timestamp { get; set; } = DateTime.Now;
    }

    public class FolderSyncBatchDto
    {
        public string SyncId { get; set; } = string.Empty;
        public int Sequence { get; set; }
        public bool Reset { get; set; }
        public bool IsFinal { get; set; }
        public List<OutlookStoreDto> Stores { get; set; } = new();
        public List<FolderDto> Folders { get; set; } = new();
    }

    public class FolderSyncCompleteDto
    {
        public string SyncId { get; set; } = string.Empty;
        public int TotalCount { get; set; }
        public bool Success { get; set; } = true;
        public string Message { get; set; } = string.Empty;
        public DateTime Timestamp { get; set; } = DateTime.Now;
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
        public string Range { get; set; } = "1m"; // 目前預期值："1d"、"1w"、"1m"。
        public int MaxCount { get; set; } = 30;
    }

    public class SearchMailsRequest
    {
        public string SearchId { get; set; } = Guid.NewGuid().ToString();
        public string StoreId { get; set; } = string.Empty;
        public List<string> ScopeFolderPaths { get; set; } = new();
        public bool IncludeSubFolders { get; set; } = true;
        public string Keyword { get; set; } = string.Empty;
        public string MatchMode { get; set; } = "contains"; // contains、exact、fuzzy、regex。fuzzy / regex 僅能在 bounded result 內後篩。
        public List<string> Fields { get; set; } = new() { "subject" };
        public DateTime? ReceivedFrom { get; set; }
        public DateTime? ReceivedTo { get; set; }
        public int MaxCount { get; set; } = 50;
        public bool IsHubSlice { get; set; }
        public string ParentCommandId { get; set; } = string.Empty;
        public int SliceIndex { get; set; }
        public int SliceCount { get; set; }
        public bool ResetSearchResults { get; set; } = true;
        public bool CompleteSearchOnSlice { get; set; } = true;
    }

    public class MailSearchSliceRequest
    {
        public string SearchId { get; set; } = string.Empty;
        public string CommandId { get; set; } = string.Empty;
        public string ParentCommandId { get; set; } = string.Empty;
        public string StoreId { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
        public DateTime? ReceivedFrom { get; set; }
        public DateTime? ReceivedTo { get; set; }
        public int MaxCount { get; set; } = 50;
        public bool IncludeBody { get; set; }
        public int SliceIndex { get; set; }
        public int SliceCount { get; set; }
        public bool ResetSearchResults { get; set; } = true;
        public bool CompleteSearchOnSlice { get; set; } = true;
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
        public List<MailItemDto> Mails { get; set; } = new();
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
        public DateTime Timestamp { get; set; } = DateTime.Now;
    }

    public class MailSearchProgressDto
    {
        public string SearchId { get; set; } = string.Empty;
        public string CommandId { get; set; } = string.Empty;
        public string Status { get; set; } = "pending"; // pending、running、completed、failed、addin_unavailable。
        public string Phase { get; set; } = string.Empty; // dispatch、store、folder、filter、completed。
        public int ProcessedStores { get; set; }
        public int TotalStores { get; set; }
        public int ProcessedFolders { get; set; }
        public int TotalFolders { get; set; }
        public int ResultCount { get; set; }
        public string CurrentStoreId { get; set; } = string.Empty;
        public string CurrentFolderPath { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
        public DateTime Timestamp { get; set; } = DateTime.Now;
        public int Percent
        {
            get
            {
                if (Status is "completed") return 100;
                if (TotalFolders > 0) return Math.Clamp((int)Math.Round(ProcessedFolders * 100.0 / TotalFolders), 0, 99);
                if (TotalStores > 0) return Math.Clamp((int)Math.Round(ProcessedStores * 100.0 / TotalStores), 0, 99);
                return Status is "running" ? 1 : 0;
            }
        }
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

    public class MailBodyDto
    {
        public string MailId { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
        public string Body { get; set; } = string.Empty;
        public string BodyHtml { get; set; } = string.Empty;
    }

    public class FetchMailAttachmentsRequest
    {
        public string MailId { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
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

    public class OpenExportedAttachmentRequest
    {
        public string ExportedAttachmentId { get; set; } = string.Empty;
    }

    public class AttachmentExportSettingsDto
    {
        public string RootPath { get; set; } = string.Empty;
        public string DefaultRootPath { get; set; } = string.Empty;
    }

    public class UpdateAttachmentExportSettingsRequest
    {
        public string RootPath { get; set; } = string.Empty;
    }

    public class MailAttachmentDto
    {
        public string MailId { get; set; } = string.Empty;
        public string Id { get; set; } = string.Empty;
        public string AttachmentId { get; set; } = string.Empty;
        public int Index { get; set; }
        public string FileName { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string ContentType { get; set; } = string.Empty;
        public long Size { get; set; }
        public bool IsExported { get; set; }
        public string ExportedAttachmentId { get; set; } = string.Empty;
        public string Path { get; set; } = string.Empty;
        public string LocalPath { get; set; } = string.Empty;
        public string FullPath { get; set; } = string.Empty;
        public string ExportedPath { get; set; } = string.Empty;
    }

    public class MailAttachmentsDto
    {
        public string MailId { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
        public List<MailAttachmentDto> Attachments { get; set; } = new();
    }

    public class ExportedMailAttachmentDto
    {
        public string MailId { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
        public string Id { get; set; } = string.Empty;
        public string AttachmentId { get; set; } = string.Empty;
        public int Index { get; set; }
        public string ExportedAttachmentId { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string ContentType { get; set; } = string.Empty;
        public long Size { get; set; }
        public string Path { get; set; } = string.Empty;
        public string LocalPath { get; set; } = string.Empty;
        public string FullPath { get; set; } = string.Empty;
        public string ExportedPath { get; set; } = string.Empty;
        public DateTime ExportedAt { get; set; } = DateTime.Now;
    }

    public class MailMarkerCommandRequest
    {
        public string MailId { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
        public string Categories { get; set; } = string.Empty;
    }

    public class MailPropertiesCommandRequest
    {
        public string MailId { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
        public bool? IsRead { get; set; }
        public string FlagInterval { get; set; } = "none"; // none、today、tomorrow、this_week、next_week、no_date、custom、complete。
        public string FlagRequest { get; set; } = string.Empty;
        public DateTime? TaskStartDate { get; set; }
        public DateTime? TaskDueDate { get; set; }
        public DateTime? TaskCompletedDate { get; set; }
        public List<string> Categories { get; set; } = new();
        public List<OutlookCategoryDto> NewCategories { get; set; } = new();
    }

    public class CategoryCommandRequest
    {
        public string Name { get; set; } = string.Empty;
        public string Color { get; set; } = string.Empty;
        public int ColorValue { get; set; }
        public string ShortcutKey { get; set; } = string.Empty;
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

    public class DeleteMailRequest
    {
        public string MailId { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
    }

    public class OutlookRuleDto
    {
        public string Name { get; set; } = string.Empty;
        public bool Enabled { get; set; }
        public int ExecutionOrder { get; set; }
        public string RuleType { get; set; } = "receive";
        public List<string> Conditions { get; set; } = new();
        public List<string> Actions { get; set; } = new();
        public List<string> Exceptions { get; set; } = new();
    }

    public class OutlookCategoryDto
    {
        public string Name { get; set; } = string.Empty;
        public string Color { get; set; } = string.Empty;
        public int ColorValue { get; set; }
        public string ShortcutKey { get; set; } = string.Empty;
    }

    public class CalendarEventDto
    {
        public string Id { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public string Location { get; set; } = string.Empty;
        public string Organizer { get; set; } = string.Empty;
        public string RequiredAttendees { get; set; } = string.Empty;
        public bool IsRecurring { get; set; }
        public string BusyStatus { get; set; } = string.Empty;
    }

    /// <summary>
    /// Hub 替 Outlook Add-in queue 的 pending command。
    /// </summary>
    public class PendingCommand
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Type { get; set; } = string.Empty; // 目前預期值："fetch_folder_roots"、"fetch_folder_children"、"fetch_mails"、"fetch_mail_body"、"fetch_mail_attachments"、"export_mail_attachment"、"fetch_rules"、"fetch_calendar"、category 與單封 mail/folder 操作。
        public FolderDiscoveryRequest? FolderDiscoveryRequest { get; set; }
        public FetchMailsRequest? MailsRequest { get; set; }
        public SearchMailsRequest? SearchMailsRequest { get; set; }
        public MailSearchSliceRequest? MailSearchSliceRequest { get; set; }
        public FetchMailBodyRequest? MailBodyRequest { get; set; }
        public FetchMailAttachmentsRequest? MailAttachmentsRequest { get; set; }
        public ExportMailAttachmentRequest? ExportMailAttachmentRequest { get; set; }
        public FetchCalendarRequest? CalendarRequest { get; set; }
        public MailMarkerCommandRequest? MailMarkerRequest { get; set; }
        public MailPropertiesCommandRequest? MailPropertiesRequest { get; set; }
        public CategoryCommandRequest? CategoryRequest { get; set; }
        public CreateFolderRequest? CreateFolderRequest { get; set; }
        public DeleteFolderRequest? DeleteFolderRequest { get; set; }
        public MoveMailRequest? MoveMailRequest { get; set; }
        public DeleteMailRequest? DeleteMailRequest { get; set; }
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
        public DateTime Timestamp { get; set; } = DateTime.Now;
    }

    public class OutlookCommandStatusDto
    {
        public string CommandId { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public string Status { get; set; } = "pending"; // pending、completed、failed、addin_unavailable。
        public bool? Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public string Payload { get; set; } = string.Empty;
        public DateTime DispatchTimestamp { get; set; } = DateTime.Now;
        public DateTime? ResultTimestamp { get; set; }
    }

    public class AddinLogEntry
    {
        public string Level { get; set; } = "info"; // "info", "warn", "error"
        public string Message { get; set; } = string.Empty;
        public DateTime Timestamp { get; set; } = DateTime.Now;
    }

    public class AddinStatusDto
    {
        public bool Connected { get; set; }
        public DateTime? LastPollTime { get; set; }
        public DateTime? LastPushTime { get; set; }
        public string LastCommand { get; set; } = string.Empty;
    }
}
