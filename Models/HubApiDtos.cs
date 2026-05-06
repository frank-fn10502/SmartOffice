namespace SmartOffice.Hub.Models
{
    // Hub HTTP API / Web UI 專用 DTO：這些型別描述 Hub 自己的 response、settings 與 diagnostics。
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

    /// <summary>
    /// Hub 接收 request-* command 後的標準回應。HTTP 200 只代表 Hub 已完成 dispatch/wait 流程；
    /// 呼叫端仍應以 commandId 查詢 command-results，並讀取對應 cached snapshot。
    /// </summary>
    public class CommandDispatchResponse
    {
        /// <summary>Hub 指派給這次 Outlook command 的 id。</summary>
        public string CommandId { get; set; } = string.Empty;
        /// <summary>目前狀態；常見值為 completed、mocked、timeout、failed、addin_unavailable、folder_cache_unavailable。</summary>
        public string Status { get; set; } = string.Empty;
        /// <summary>Hub 或 Outlook AddIn 回報的簡短訊息。</summary>
        public string Message { get; set; } = string.Empty;
    }

    /// <summary>
    /// Mail search dispatch 的標準回應；searchId 用於查詢搜尋進度與 cached search results。
    /// </summary>
    public class MailSearchDispatchResponse : CommandDispatchResponse
    {
        /// <summary>搜尋 correlation id；可用於 GET /api/outlook/mail-search/progress/{searchId}。</summary>
        public string SearchId { get; set; } = string.Empty;
        /// <summary>Hub 展開 folder scope 後 dispatch 給 AddIn 的 folder slice 數量。</summary>
        public int SliceCount { get; set; }
    }

    /// <summary>
    /// Folder roots dispatch 完成後的標準回應，附帶目前 cached folder snapshot 計數。
    /// </summary>
    public class FolderRequestDispatchResponse : CommandDispatchResponse
    {
        public int Stores { get; set; }
        public int Folders { get; set; }
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
        public string Status { get; set; } = "pending"; // pending、completed、failed、addin_unavailable。
        public bool? Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public string Payload { get; set; } = string.Empty;
        public DateTime DispatchTimestamp { get; set; } = DateTime.Now;
        public DateTime? ResultTimestamp { get; set; }
    }

    public class AddinStatusDto
    {
        public bool Connected { get; set; }
        public DateTime? LastPollTime { get; set; }
        public DateTime? LastPushTime { get; set; }
        public string LastCommand { get; set; } = string.Empty;
    }
}
