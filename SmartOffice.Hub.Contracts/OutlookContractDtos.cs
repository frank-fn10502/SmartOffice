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
        public string MessageClass { get; set; } = string.Empty;
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
        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
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
        public string AccountSmtpAddress { get; set; } = string.Empty;
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
        public DateTime ExportedAt { get; set; } = DateTime.UtcNow;
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
        public bool SmartOfficeOwned { get; set; }
        public string SmartOfficeEventId { get; set; } = string.Empty;
    }

    public class CalendarRoomDto
    {
        public string Id { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string SmtpAddress { get; set; } = string.Empty;
        public string RawAddress { get; set; } = string.Empty;
        public string Source { get; set; } = string.Empty;
    }

    public class AddressBookContactDto
    {
        public string Id { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string SmtpAddress { get; set; } = string.Empty;
        public string RawAddress { get; set; } = string.Empty;
        public string AddressType { get; set; } = string.Empty;
        public string EntryUserType { get; set; } = string.Empty;
        public string Source { get; set; } = string.Empty;
        public string CompanyName { get; set; } = string.Empty;
        public string JobTitle { get; set; } = string.Empty;
        public string Department { get; set; } = string.Empty;
        public string OfficeLocation { get; set; } = string.Empty;
        public string BusinessTelephoneNumber { get; set; } = string.Empty;
        public string MobileTelephoneNumber { get; set; } = string.Empty;
        public string Domain { get; set; } = string.Empty;
        public bool IsKnown { get; set; }
        public bool IsLikelySelf { get; set; }
        public bool IsRelatedToSelf { get; set; }
        public bool IsGroup { get; set; }
        public int MemberCount { get; set; }
        public bool GroupMembersLoaded { get; set; }
        public bool GroupMembersLoading { get; set; }
        public string GroupMembersRequestId { get; set; } = string.Empty;
        public DateTime? GroupMembersUpdatedAt { get; set; }
        public int RelationScore { get; set; }
        public int MailCount { get; set; }
        public int CalendarCount { get; set; }
        public int SenderCount { get; set; }
        public int RecipientCount { get; set; }
        public int OrganizerCount { get; set; }
        public int AttendeeCount { get; set; }
        public int GroupMemberCount { get; set; }
        public DateTime? FirstSeen { get; set; }
        public DateTime? LastSeen { get; set; }
        public List<string> RelationKinds { get; set; } = new List<string>();
        public List<string> Sources { get; set; } = new List<string>();
        public List<string> MemberSmtpAddresses { get; set; } = new List<string>();
        public List<string> MemberGroupSmtpAddresses { get; set; } = new List<string>();
        public List<string> MemberOfGroupSmtpAddresses { get; set; } = new List<string>();
        public List<string> FolderPaths { get; set; } = new List<string>();
        public List<string> RecentMailIds { get; set; } = new List<string>();
        public List<string> SampleSubjects { get; set; } = new List<string>();
    }

    public class AddressBookSyncRequest
    {
        public bool IncludeOutlookContacts { get; set; } = true;
        public bool IncludeAddressLists { get; set; } = true;
        public int MaxContacts { get; set; } = 5000;
        public int MaxAddressEntriesPerList { get; set; } = 2000;
        public int MaxGroupMembers { get; set; } = 0;
        public int MaxGroupDepth { get; set; } = 1;
        public string AddressListId { get; set; } = string.Empty;
        public string AddressListName { get; set; } = string.Empty;
        public int Offset { get; set; }
        public int PageSize { get; set; } = 100;
        public string GroupId { get; set; } = string.Empty;
        public string GroupSmtpAddress { get; set; } = string.Empty;
        public bool ForceRefresh { get; set; }
    }

    public class AddressBookBatchDto
    {
        public string BatchId { get; set; } = string.Empty;
        public int Sequence { get; set; }
        public bool Reset { get; set; }
        public bool IsFinal { get; set; }
        public int TotalCount { get; set; }
        public List<AddressBookContactDto> Contacts { get; set; } = new List<AddressBookContactDto>();
    }

    public class AddressBookRootDto
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string AddressListType { get; set; } = string.Empty;
        public string Source { get; set; } = string.Empty;
        public int EntryCount { get; set; }
        public bool CanPageEntries { get; set; } = true;
    }

    public class AddressBookRootsBatchDto
    {
        public string RequestId { get; set; } = string.Empty;
        public List<AddressBookRootDto> Roots { get; set; } = new List<AddressBookRootDto>();
    }

    public class AddressBookListEntriesRequest
    {
        public string AddressListId { get; set; } = string.Empty;
        public string AddressListName { get; set; } = string.Empty;
        public int Offset { get; set; }
        public int PageSize { get; set; } = 100;
    }

    public class AddressBookListEntriesPageDto
    {
        public string RequestId { get; set; } = string.Empty;
        public string AddressListId { get; set; } = string.Empty;
        public string AddressListName { get; set; } = string.Empty;
        public int Offset { get; set; }
        public int PageSize { get; set; }
        public int TotalCount { get; set; }
        public bool HasMore { get; set; }
        public List<AddressBookContactDto> Contacts { get; set; } = new List<AddressBookContactDto>();
    }

    public class AddressBookGroupMembersRequest
    {
        public string GroupId { get; set; } = string.Empty;
        public string GroupSmtpAddress { get; set; } = string.Empty;
        public int MaxMembers { get; set; } = 0;
        public bool ForceRefresh { get; set; }
    }

    public class AddressBookGroupMembersBatchDto
    {
        public string GroupId { get; set; } = string.Empty;
        public string GroupSmtpAddress { get; set; } = string.Empty;
        public string BatchId { get; set; } = string.Empty;
        public int Sequence { get; set; }
        public bool Reset { get; set; }
        public bool IsFinal { get; set; }
        public int TotalCount { get; set; }
        public List<AddressBookContactDto> Members { get; set; } = new List<AddressBookContactDto>();
    }

    public class AddressBookRelationLookupRequest
    {
        public string Query { get; set; } = string.Empty;
        public string TargetKind { get; set; } = string.Empty;
        public string Id { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string SmtpAddress { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public string GroupId { get; set; } = string.Empty;
        public string GroupSmtpAddress { get; set; } = string.Empty;
        public int Take { get; set; } = 50;
    }

    public class AddressBookRelationLookupResponse
    {
        public string Query { get; set; } = string.Empty;
        public string TargetKind { get; set; } = string.Empty;
        public string State { get; set; } = "unknown";
        public string Message { get; set; } = string.Empty;
        public AddressBookContactDto Target { get; set; }
        public List<AddressBookContactDto> Matches { get; set; } = new List<AddressBookContactDto>();
        public List<AddressBookContactDto> Members { get; set; } = new List<AddressBookContactDto>();
        public List<AddressBookContactDto> MemberGroups { get; set; } = new List<AddressBookContactDto>();
        public List<AddressBookContactDto> MemberOfGroups { get; set; } = new List<AddressBookContactDto>();
        public List<AddressBookContactDto> ContainingGroups { get; set; } = new List<AddressBookContactDto>();
        public bool IsGroup { get; set; }
        public bool IsLikelySelf { get; set; }
        public bool IsRelatedToSelf { get; set; }
        public bool GroupMembersLoaded { get; set; }
        public bool GroupMembersLoading { get; set; }
        public AddressBookRecipientRelevanceDto RecipientRelevance { get; set; } = new AddressBookRecipientRelevanceDto();
    }

    public class AddressBookRecipientRelevanceDto
    {
        public int Score { get; set; }
        public string Level { get; set; } = "unknown";
        public string Summary { get; set; } = string.Empty;
        public int RouteDepth { get; set; }
        public int DirectPersonCount { get; set; }
        public int DirectGroupCount { get; set; }
        public int AudienceSize { get; set; }
        public bool IncludesSelf { get; set; }
        public bool IncludesSelfDirectly { get; set; }
        public List<string> Reasons { get; set; } = new List<string>();
    }

    public class OutlookProfileMailStatsDto
    {
        public int LoadedCount { get; set; }
        public int UnreadCount { get; set; }
        public int AttachmentMailCount { get; set; }
    }

    public class OutlookProfileDto
    {
        public string State { get; set; } = "ready";
        public string Message { get; set; } = string.Empty;
        public string MailboxName { get; set; } = string.Empty;
        public string SmtpAddress { get; set; } = string.Empty;
        public AddressBookContactDto SelfContact { get; set; }
        public List<OutlookProfileGroupNodeDto> GroupTree { get; set; } = new List<OutlookProfileGroupNodeDto>();
        public List<AddressBookContactDto> Groups { get; set; } = new List<AddressBookContactDto>();
        public List<AddressBookContactDto> SameGroupPeople { get; set; } = new List<AddressBookContactDto>();
        public List<OutlookStoreDto> OstStores { get; set; } = new List<OutlookStoreDto>();
        public List<OutlookStoreDto> PstStores { get; set; } = new List<OutlookStoreDto>();
        public List<OutlookStoreDto> OtherStores { get; set; } = new List<OutlookStoreDto>();
        public List<OutlookStoreDto> Stores { get; set; } = new List<OutlookStoreDto>();
        public OutlookProfileMailStatsDto MailStats { get; set; } = new OutlookProfileMailStatsDto();
    }

    public class OutlookProfileGroupNodeDto
    {
        public AddressBookContactDto Contact { get; set; }
        public List<OutlookProfileGroupNodeDto> Children { get; set; } = new List<OutlookProfileGroupNodeDto>();
    }

    public class AddinLogEntry
    {
        public string Level { get; set; } = "info";
        public string Message { get; set; } = string.Empty;
        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
    }
}
