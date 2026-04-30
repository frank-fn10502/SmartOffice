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
        public string FolderPath { get; set; } = string.Empty;
        public int ItemCount { get; set; }
        public List<FolderDto> SubFolders { get; set; } = new();
    }

    public class FetchMailsRequest
    {
        public string FolderPath { get; set; } = string.Empty;
        public string Range { get; set; } = "1d"; // 目前預期值："1d"、"1w"、"1m"。
        public int MaxCount { get; set; } = 10;
    }

    public class FetchCalendarRequest
    {
        public int DaysForward { get; set; } = 14;
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
        public string Type { get; set; } = string.Empty; // 目前預期值："fetch_mails"、"fetch_folders"、"fetch_rules"、"fetch_calendar"、category 與單封 mail/folder 操作。
        public FetchMailsRequest? MailsRequest { get; set; }
        public FetchCalendarRequest? CalendarRequest { get; set; }
        public MailMarkerCommandRequest? MailMarkerRequest { get; set; }
        public MailPropertiesCommandRequest? MailPropertiesRequest { get; set; }
        public CategoryCommandRequest? CategoryRequest { get; set; }
        public CreateFolderRequest? CreateFolderRequest { get; set; }
        public DeleteFolderRequest? DeleteFolderRequest { get; set; }
        public MoveMailRequest? MoveMailRequest { get; set; }
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
