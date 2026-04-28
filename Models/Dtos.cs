namespace SmartOffice.Hub.Models
{
    public class MailItemDto
    {
        public string Subject { get; set; } = string.Empty;
        public string SenderName { get; set; } = string.Empty;
        public string SenderEmail { get; set; } = string.Empty;
        public DateTime ReceivedTime { get; set; }
        public string Body { get; set; } = string.Empty;
        public string BodyHtml { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
    }

    public class ChatMessageDto
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Source { get; set; } = "outlook"; // Expected values today: "outlook" or "web".
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
        public string Range { get; set; } = "1d"; // Expected values today: "1d", "1w", "1m".
        public int MaxCount { get; set; } = 10;
    }

    /// <summary>
    /// Pending command queued by Hub for Outlook Add-in to pick up.
    /// </summary>
    public class PendingCommand
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Type { get; set; } = string.Empty; // Expected values today: "fetch_mails", "fetch_folders".
        public FetchMailsRequest? MailsRequest { get; set; }
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
