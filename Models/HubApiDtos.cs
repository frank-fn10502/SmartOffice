namespace SmartOffice.Hub.Models
{
    // Hub HTTP API / Web UI 專用 DTO：這些型別描述 Hub 自己的 response、settings 與 diagnostics。
    public class MailSearchProgressDto
    {
        public string SearchId { get; set; } = string.Empty;
        public string CommandId { get; set; } = string.Empty;
        public string Status { get; set; } = "pending"; // pending、running、completed、failed、outlook_unavailable。
        public string Phase { get; set; } = string.Empty; // dispatch、store、folder、filter、completed。
        public int ProcessedStores { get; set; }
        public int TotalStores { get; set; }
        public int ProcessedFolders { get; set; }
        public int TotalFolders { get; set; }
        public int ResultCount { get; set; }
        public string CurrentStoreId { get; set; } = string.Empty;
        public string CurrentFolderPath { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
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

    public class OpenExportedAttachmentRequest
    {
        public string ExportedAttachmentId { get; set; } = string.Empty;
    }

    public class RequestMailsApiRequest
    {
        public string FolderPath { get; set; } = string.Empty;
        public double? LookbackHours { get; set; }
        public DateTime? ReceivedFrom { get; set; }
        public DateTime? ReceivedTo { get; set; }
        public int MaxCount { get; set; } = 30;
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

    public class OutlookRequestResponse
    {
        public string RequestId { get; set; } = string.Empty;
        public string Request { get; set; } = string.Empty;
        public string State { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
        public object Data { get; set; } = new { };
    }

    public class FetchResultRequest
    {
        public string RequestId { get; set; } = string.Empty;
        public string Cursor { get; set; } = string.Empty;
        public int Take { get; set; } = 100;
    }

    public class FetchResultNext
    {
        public string Cursor { get; set; } = string.Empty;
        public bool HasMore { get; set; }
    }

    public class FetchResultResponse
    {
        public string RequestId { get; set; } = string.Empty;
        public string Request { get; set; } = string.Empty;
        public string State { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
        public FetchResultNext Next { get; set; } = new();
        public object Data { get; set; } = new { };
    }

    public class FolderMailsRequest
    {
        public string FolderPath { get; set; } = string.Empty;
        public bool IncludeSubFolders { get; set; } = true;
        public DateTime? ReceivedFrom { get; set; }
        public DateTime? ReceivedTo { get; set; }
        public int MaxCount { get; set; } = 30;
    }

    public class OpenExportedAttachmentResponse
    {
        public string Status { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
    }

    public class OutlookCommandStatusDto
    {
        public string CommandId { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public string Status { get; set; } = "pending"; // pending、completed、failed、outlook_unavailable。
        public bool? Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public string Payload { get; set; } = string.Empty;
        public DateTime DispatchTimestamp { get; set; } = DateTime.UtcNow;
        public DateTime? ResultTimestamp { get; set; }
    }

    public class AddinStatusDto
    {
        public bool Connected { get; set; }
        public DateTime? LastPollTime { get; set; }
        public DateTime? LastPushTime { get; set; }
        public string LastCommand { get; set; } = string.Empty;
    }

    public class AddressBookResponse
    {
        public string Query { get; set; } = string.Empty;
        public int TotalCount { get; set; }
        public List<AddressBookContactDto> Contacts { get; set; } = new();
    }

    public class AddressBookLookupResponse
    {
        public string Query { get; set; } = string.Empty;
        public string State { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
        public AddressBookContactDto? Contact { get; set; }
        public List<AddressBookContactDto> Suggestions { get; set; } = new();
    }

    public class AddressBookGroupMembersResponse
    {
        public string State { get; set; } = "not_loaded";
        public string Message { get; set; } = string.Empty;
        public string GroupKey { get; set; } = string.Empty;
        public string GroupSmtpAddress { get; set; } = string.Empty;
        public string RequestId { get; set; } = string.Empty;
        public int TotalCount { get; set; }
        public DateTime? UpdatedAt { get; set; }
        public List<AddressBookContactDto> Members { get; set; } = new();
    }

    public class AddressBookMergeSuggestionRequest
    {
        public List<string> Recipients { get; set; } = new();
    }

    public class AddressBookMergeSuggestionDto
    {
        public string GroupSmtpAddress { get; set; } = string.Empty;
        public string GroupDisplayName { get; set; } = string.Empty;
        public List<AddressBookContactDto> CoveredContacts { get; set; } = new();
        public List<string> CoveredRecipientKeys { get; set; } = new();
        public string Message { get; set; } = string.Empty;
    }

    public class AddressBookMergeSuggestionResponse
    {
        public string State { get; set; } = "ok";
        public List<AddressBookMergeSuggestionDto> Suggestions { get; set; } = new();
    }
}
