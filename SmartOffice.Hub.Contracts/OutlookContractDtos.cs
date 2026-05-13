using System;
using System.Collections.Generic;

namespace SmartOffice.Hub.Contracts
{
    public class MailItemDto
    {
        public string Id { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public OutlookRecipientDto Sender { get; set; } = new OutlookRecipientDto();
        public List<OutlookRecipientDto> ToRecipients { get; set; } = new List<OutlookRecipientDto>();
        public List<OutlookRecipientDto> CcRecipients { get; set; } = new List<OutlookRecipientDto>();
        public List<OutlookRecipientDto> BccRecipients { get; set; } = new List<OutlookRecipientDto>();
        public DateTime ReceivedTime { get; set; }
        public string Body { get; set; } = string.Empty;
        public string BodyHtml { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
        public string ConversationId { get; set; } = string.Empty;
        public string ConversationTopic { get; set; } = string.Empty;
        public string ConversationIndex { get; set; } = string.Empty;
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

    public class OutlookRecipientDto
    {
        public string RecipientKind { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string SmtpAddress { get; set; } = string.Empty;
        public string RawAddress { get; set; } = string.Empty;
        public string AddressType { get; set; } = string.Empty;
        public string EntryUserType { get; set; } = string.Empty;
        public bool IsGroup { get; set; }
        public bool IsResolved { get; set; }
        public List<OutlookRecipientDto> Members { get; set; } = new List<OutlookRecipientDto>();
    }

    public class ChatMessageDto
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Source { get; set; } = string.Empty;
        public string Text { get; set; } = string.Empty;
        public DateTime Timestamp { get; set; } = DateTime.Now;
    }

    public enum OutlookFolderType
    {
        Unknown = 0,
        StoreRoot,
        Mail,
        Inbox,
        Sent,
        Drafts,
        Deleted,
        Junk,
        Archive,
        Outbox,
        SyncIssues,
        Conflicts,
        LocalFailures,
        ServerFailures,
        Calendar,
        Contacts,
        Tasks,
        Notes,
        Journal,
        RssFeeds,
        ConversationHistory,
        ConversationActionSettings,
        OtherSystem,
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
        public OutlookFolderType FolderType { get; set; } = OutlookFolderType.Unknown;
        public int DefaultItemType { get; set; } = -1;
        public bool IsHidden { get; set; }
        public bool IsSystem { get; set; }
        public bool HasChildren { get; set; }
        public bool ChildrenLoaded { get; set; }
        public string DiscoveryState { get; set; } = "partial";
    }

    public class OutlookStoreDto
    {
        public string StoreId { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string StoreKind { get; set; } = string.Empty;
        public string StoreFilePath { get; set; } = string.Empty;
        public string RootFolderPath { get; set; } = string.Empty;
    }

    public class FolderSnapshotDto
    {
        public List<OutlookStoreDto> Stores { get; set; } = new List<OutlookStoreDto>();
        public List<FolderDto> Folders { get; set; } = new List<FolderDto>();
    }

    public class MailBodyDto
    {
        public string MailId { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
        public string Body { get; set; } = string.Empty;
        public string BodyHtml { get; set; } = string.Empty;
    }

    public class MailConversationDto
    {
        public string MailId { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
        public string ConversationId { get; set; } = string.Empty;
        public string ConversationTopic { get; set; } = string.Empty;
        public List<MailItemDto> Mails { get; set; } = new List<MailItemDto>();
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
        public List<MailAttachmentDto> Attachments { get; set; } = new List<MailAttachmentDto>();
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

    public class OutlookRuleDto
    {
        public string StoreId { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public bool Enabled { get; set; }
        public int ExecutionOrder { get; set; }
        public string RuleType { get; set; } = "receive";
        public bool IsLocalRule { get; set; }
        public bool CanModifyDefinition { get; set; } = true;
        public List<string> Conditions { get; set; } = new List<string>();
        public List<string> Actions { get; set; } = new List<string>();
        public List<string> Exceptions { get; set; } = new List<string>();
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
        public OutlookRecipientDto Organizer { get; set; } = new OutlookRecipientDto();
        public List<OutlookRecipientDto> RequiredAttendees { get; set; } = new List<OutlookRecipientDto>();
        public bool IsRecurring { get; set; }
        public string BusyStatus { get; set; } = string.Empty;
    }

    public class AddinLogEntry
    {
        public string Level { get; set; } = "info";
        public string Message { get; set; } = string.Empty;
        public DateTime Timestamp { get; set; } = DateTime.Now;
    }
}
