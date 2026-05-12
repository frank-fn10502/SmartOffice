namespace SmartOffice.Hub.Models
{
    // AddIn contract DTO：Hub 透過 SignalR dispatch command 給 AddIn，AddIn 也用這些 DTO 回推結果。
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
        public DateTime? ReceivedFrom { get; set; }
        public DateTime? ReceivedTo { get; set; }
        public int MaxCount { get; set; } = 30;
    }

    /// <summary>
    /// Web UI / AI 要求 Hub 以 Outlook 內建搜尋查詢郵件的 request。
    /// Hub 會將此 request 展開成 AddIn 實際處理的 MailSearchSliceRequest。
    /// 直接列出 folder mails 請使用 FolderMailsRequest，不走 search。
    /// </summary>
    public class SearchMailsRequest
    {
        /// <summary>搜尋 correlation id；呼叫端可自訂，未提供時 Hub 會產生。</summary>
        public string SearchId { get; set; } = Guid.NewGuid().ToString();
        /// <summary>限制在指定 Outlook store 搜尋；空字串代表全部 store。</summary>
        public string StoreId { get; set; } = string.Empty;
        /// <summary>限制搜尋範圍的 folder path；空陣列只允許搭配 storeId 或 allowGlobalScope 使用。</summary>
        public List<string> ScopeFolderPaths { get; set; } = new();
        /// <summary>允許空 storeId 與空 scopeFolderPaths 時搜尋全部已載入的可搜尋 mail folder。</summary>
        public bool AllowGlobalScope { get; set; }
        /// <summary>ScopeFolderPaths 有值時是否包含子資料夾。</summary>
        public bool IncludeSubFolders { get; set; } = true;
        /// <summary>文字搜尋關鍵字；空白時只套用其他篩選條件。</summary>
        public string Keyword { get; set; } = string.Empty;
        /// <summary>keyword 套用欄位；正式值為 subject、sender、body，預設 subject。</summary>
        public List<string> TextFields { get; set; } = new() { "subject" };
        /// <summary>Outlook category 篩選；任一分類符合即可。</summary>
        public List<string> CategoryNames { get; set; } = new();
        /// <summary>附件篩選；true 表示包含附件，false 表示不含附件，null 表示不限。</summary>
        public bool? HasAttachments { get; set; }
        /// <summary>旗標篩選；正式值為 any、flagged、unflagged。</summary>
        public string FlagState { get; set; } = "any";
        /// <summary>已讀篩選；正式值為 any、unread、read。</summary>
        public string ReadState { get; set; } = "any";
        /// <summary>收到時間起點，可獨立使用。</summary>
        public DateTime? ReceivedFrom { get; set; }
        /// <summary>收到時間終點，可獨立使用。</summary>
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
        public List<string> TextFields { get; set; } = new() { "subject" };
        public List<string> CategoryNames { get; set; } = new();
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
        public List<MailItemDto> Mails { get; set; } = new();
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
        public DateTime Timestamp { get; set; } = DateTime.Now;
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

    public class OutlookRuleConditionsRequest
    {
        public List<string> SubjectContains { get; set; } = new();
        public List<string> BodyContains { get; set; } = new();
        public List<string> SenderAddressContains { get; set; } = new();
        public List<string> Categories { get; set; } = new();
        public bool? HasAttachment { get; set; }
    }

    public class OutlookRuleActionsRequest
    {
        public string MoveToFolderPath { get; set; } = string.Empty;
        public List<string> AssignCategories { get; set; } = new();
        public bool MarkAsTask { get; set; }
        public bool StopProcessingMoreRules { get; set; } = true;
    }

    public class OutlookRuleCommandRequest
    {
        public string Operation { get; set; } = "upsert"; // upsert、delete、set_enabled。
        public string StoreId { get; set; } = string.Empty;
        public string RuleName { get; set; } = string.Empty;
        public string OriginalRuleName { get; set; } = string.Empty;
        public int? OriginalExecutionOrder { get; set; }
        public string RuleType { get; set; } = "receive";
        public bool Enabled { get; set; } = true;
        public int? ExecutionOrder { get; set; }
        public OutlookRuleConditionsRequest Conditions { get; set; } = new();
        public OutlookRuleActionsRequest Actions { get; set; } = new();
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
        public List<string> MailIds { get; set; } = new();
        public string SourceFolderPath { get; set; } = string.Empty;
        public List<string> SourceFolderPaths { get; set; } = new();
        public string DestinationFolderPath { get; set; } = string.Empty;
        public bool ContinueOnError { get; set; } = true;
    }

    public class DeleteMailRequest
    {
        public string MailId { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
    }

    /// <summary>
    /// Hub 替 Outlook AddIn queue 的 pending command。
    /// </summary>
    public class PendingCommand
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Type { get; set; } = string.Empty; // 目前預期值："fetch_folder_roots"、"fetch_folder_children"、"fetch_mails"、"fetch_mail_body"、"fetch_mail_attachments"、"fetch_mail_conversation"、"export_mail_attachment"、"fetch_rules"、"fetch_calendar"、category 與單封 mail/folder 操作。
        public FolderDiscoveryRequest? FolderDiscoveryRequest { get; set; }
        public FindFolderRequest? FindFolderRequest { get; set; }
        public FetchMailsRequest? MailsRequest { get; set; }
        public SearchMailsRequest? SearchMailsRequest { get; set; }
        public MailSearchSliceRequest? MailSearchSliceRequest { get; set; }
        public FolderMailsSliceRequest? FolderMailsSliceRequest { get; set; }
        public FetchMailBodyRequest? MailBodyRequest { get; set; }
        public FetchMailAttachmentsRequest? MailAttachmentsRequest { get; set; }
        public FetchMailConversationRequest? MailConversationRequest { get; set; }
        public ExportMailAttachmentRequest? ExportMailAttachmentRequest { get; set; }
        public FetchCalendarRequest? CalendarRequest { get; set; }
        public MailPropertiesCommandRequest? MailPropertiesRequest { get; set; }
        public CategoryCommandRequest? CategoryRequest { get; set; }
        public OutlookRuleCommandRequest? RuleRequest { get; set; }
        public CreateFolderRequest? CreateFolderRequest { get; set; }
        public DeleteFolderRequest? DeleteFolderRequest { get; set; }
        public MoveMailRequest? MoveMailRequest { get; set; }
        public MoveMailsRequest? MoveMailsRequest { get; set; }
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
}
