using Microsoft.AspNetCore.SignalR;
using SmartOffice.Hub.Models;

namespace SmartOffice.Hub.Services
{
    public class MailStore
    {
        private readonly object _lock = new();
        private static readonly HashSet<string> SystemFolderNames = new(StringComparer.OrdinalIgnoreCase)
        {
            "Sync Issues",
            "Conflicts",
            "Local Failures",
            "Server Failures",
            "RSS Feeds",
            "RSS Subscriptions",
            "Quick Step Settings",
            "Conversation Action Settings",
            "Conversation History",
            "Social Activity Notifications",
            "ExternalContacts",
            "MyContactsExtended",
            "Recipient Cache",
            "PersonMetadata",
            "{A9E2BC46-B3A0-4243-B315-60D991004455}",
            "{06967759-274D-40B2-A3EB-D7F9E73727D7}",
            "Yammer Root",
            "Files",
            "GraphFilesAndWorkPagesFolder",
            "Finder",
            "Common Views",
            "Reminders",
            "Shortcuts",
            "Spooler Queue",
            "公用資料夾",
            "公用文件夾",
            "Public Folders",
            "?ffentliche Ordner",
            "Dossiers publics",
            "Recoverable Items",
            "Deletions",
            "Purges",
            "Versions",
            "DiscoveryHolds",
            "Calendar Logging",
            "Audits",
            "AdminAuditLogs",
            "FreeBusy Data",
            "Top of Information Store",
            "System",
            "ExchangeSyncData",
            "AllItems",
            "AllContacts",
            "Freebusy Data",
            "Schedule",
            "GAL Contacts",
            "OAB Version 2",
            "OAB Version 3",
            "OAB Version 4",
            "Offline Address Book",
        };
        private static readonly string[] SystemFolderNameFragments =
        {
            "公用資料夾",
            "公用文件夾",
            "Public Folders",
            "Dossiers publics",
            "ffentliche Ordner",
        };
        private static readonly HashSet<OutlookFolderType> BlockedFolderTypes = new()
        {
            OutlookFolderType.SyncIssues,
            OutlookFolderType.Conflicts,
            OutlookFolderType.LocalFailures,
            OutlookFolderType.ServerFailures,
            OutlookFolderType.Calendar,
            OutlookFolderType.Contacts,
            OutlookFolderType.Tasks,
            OutlookFolderType.Notes,
            OutlookFolderType.Journal,
            OutlookFolderType.RssFeeds,
            OutlookFolderType.ConversationHistory,
            OutlookFolderType.ConversationActionSettings,
            OutlookFolderType.OtherSystem,
        };
        private List<MailItemDto> _mails = new();
        private List<FolderDto> _folders = new();
        private List<OutlookStoreDto> _stores = new();
        private List<OutlookRuleDto> _rules = new();
        private List<OutlookCategoryDto> _categories = new();
        private List<CalendarEventDto> _calendarEvents = new();
        private List<MailItemDto> _mailSearchResults = new();
        private List<MailItemDto> _folderMailResults = new();
        private DateTime _folderCacheUpdatedAt = DateTime.MinValue;
        private readonly Dictionary<string, List<MailItemDto>> _mailSearchResultsBySearchId = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, List<MailItemDto>> _folderMailResultsByRequestId = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, MailSearchProgressDto> _mailSearchProgress = new();
        private readonly Dictionary<string, SearchMailsRequest> _mailSearchRequests = new();
        private readonly Dictionary<string, MailAttachmentsDto> _attachments = new();
        private readonly Dictionary<string, MailConversationDto> _conversations = new();
        private readonly Dictionary<string, ExportedMailAttachmentDto> _exportedAttachments = new();

        public void SetMails(List<MailItemDto> mails)
        {
            lock (_lock) { _mails = new List<MailItemDto>(mails); }
        }

        public void UpsertMail(MailItemDto mail)
        {
            lock (_lock)
            {
                UpsertKnownMail(_mails, mail);
                UpsertKnownMail(_mailSearchResults, mail);
                UpsertKnownMail(_folderMailResults, mail);
                foreach (var results in _mailSearchResultsBySearchId.Values)
                    UpsertKnownMail(results, mail);
                foreach (var results in _folderMailResultsByRequestId.Values)
                    UpsertKnownMail(results, mail);
            }
        }

        public void UpdateMailBody(MailBodyDto body)
        {
            lock (_lock)
            {
                UpdateMailBody(_mails, body);
                UpdateMailBody(_mailSearchResults, body);
                UpdateMailBody(_folderMailResults, body);
                foreach (var results in _mailSearchResultsBySearchId.Values)
                    UpdateMailBody(results, body);
                foreach (var results in _folderMailResultsByRequestId.Values)
                    UpdateMailBody(results, body);
            }
        }

        public void SetMailAttachments(MailAttachmentsDto attachments)
        {
            lock (_lock)
            {
                NormalizeMailAttachments(attachments);

                var exported = attachments.Attachments
                    .Select(item => _exportedAttachments.Values.FirstOrDefault(exported =>
                        exported.MailId == item.MailId && exported.AttachmentId == item.AttachmentId))
                    .Where(item => item is not null)
                    .ToDictionary(item => item!.AttachmentId, item => item!);

                foreach (var attachment in attachments.Attachments)
                {
                    if (!exported.TryGetValue(attachment.AttachmentId, out var exportedAttachment)) continue;
                    attachment.IsExported = true;
                    attachment.ExportedAttachmentId = exportedAttachment.ExportedAttachmentId;
                    attachment.ExportedPath = exportedAttachment.ExportedPath;
                }

                _attachments[attachments.MailId] = CloneMailAttachments(attachments);
            }
        }

        public MailAttachmentsDto? GetMailAttachments(string mailId)
        {
            lock (_lock)
            {
                return _attachments.TryGetValue(mailId, out var attachments)
                    ? CloneMailAttachments(attachments)
                    : null;
            }
        }

        public void SetMailConversation(MailConversationDto conversation)
        {
            lock (_lock)
            {
                NormalizeMailConversation(conversation);
                _conversations[conversation.MailId] = CloneMailConversation(conversation);
            }
        }

        public MailConversationDto? GetMailConversation(string mailId)
        {
            lock (_lock)
            {
                return _conversations.TryGetValue(mailId, out var conversation)
                    ? CloneMailConversation(conversation)
                    : null;
            }
        }

        public void UpsertExportedAttachment(ExportedMailAttachmentDto attachment)
        {
            lock (_lock)
            {
                NormalizeExportedAttachment(attachment);

                _exportedAttachments[attachment.ExportedAttachmentId] = CloneExportedAttachment(attachment);

                if (!_attachments.TryGetValue(attachment.MailId, out var attachments)) return;
                var item = attachments.Attachments.FirstOrDefault(next => next.AttachmentId == attachment.AttachmentId);
                if (item is null) return;
                item.IsExported = true;
                item.ExportedAttachmentId = attachment.ExportedAttachmentId;
                item.ExportedPath = attachment.ExportedPath;
            }
        }

        public bool TryGetExportedAttachment(string exportedAttachmentId, out ExportedMailAttachmentDto attachment)
        {
            lock (_lock)
            {
                if (_exportedAttachments.TryGetValue(exportedAttachmentId, out var found))
                {
                    attachment = CloneExportedAttachment(found);
                    return true;
                }
            }

            attachment = new ExportedMailAttachmentDto();
            return false;
        }

        public List<MailItemDto> GetMails()
        {
            lock (_lock) { return _mails.Select(CloneMail).ToList(); }
        }

        public void BeginMailSearch(bool reset = true, string searchId = "")
        {
            if (!reset) return;
            lock (_lock)
            {
                _mailSearchResults = new List<MailItemDto>();
                if (!string.IsNullOrWhiteSpace(searchId))
                    _mailSearchResultsBySearchId[searchId] = new List<MailItemDto>();
            }
        }

        public void BeginFolderMails(bool reset = true, string folderMailsId = "")
        {
            if (!reset) return;
            lock (_lock)
            {
                _folderMailResults = new List<MailItemDto>();
                if (!string.IsNullOrWhiteSpace(folderMailsId))
                    _folderMailResultsByRequestId[folderMailsId] = new List<MailItemDto>();
            }
        }

        public MailSearchProgressDto StartMailSearchProgress(SearchMailsRequest request, string commandId)
        {
            lock (_lock)
            {
                var progress = new MailSearchProgressDto
                {
                    SearchId = request.SearchId,
                    CommandId = commandId,
                    Status = "pending",
                    Phase = "dispatch",
                    TotalFolders = request.ScopeFolderPaths.Count(path => !string.IsNullOrWhiteSpace(path)),
                    TotalStores = string.IsNullOrWhiteSpace(request.StoreId) ? 0 : 1,
                    Message = "Mail search dispatched to Outlook AddIn.",
                    Timestamp = DateTime.Now,
                };
                _mailSearchRequests[request.SearchId] = CloneSearchMailsRequest(request);
                _mailSearchProgress[progress.SearchId] = CloneMailSearchProgress(progress);
                return CloneMailSearchProgress(progress);
            }
        }

        public MailSearchProgressDto UpdateMailSearchProgress(MailSearchProgressDto progress)
        {
            lock (_lock)
            {
                if (_mailSearchProgress.TryGetValue(progress.SearchId, out var current))
                {
                    if (string.IsNullOrWhiteSpace(progress.CommandId)) progress.CommandId = current.CommandId;
                    if (string.IsNullOrWhiteSpace(progress.Status)) progress.Status = current.Status;
                    if (progress.TotalFolders <= 0) progress.TotalFolders = current.TotalFolders;
                    if (progress.TotalStores <= 0) progress.TotalStores = current.TotalStores;
                }

                progress.Timestamp = progress.Timestamp == default ? DateTime.Now : progress.Timestamp;
                _mailSearchProgress[progress.SearchId] = CloneMailSearchProgress(progress);
                return CloneMailSearchProgress(progress);
            }
        }

        public MailSearchProgressDto? GetMailSearchProgress(string searchId)
        {
            lock (_lock)
            {
                return _mailSearchProgress.TryGetValue(searchId, out var progress)
                    ? CloneMailSearchProgress(progress)
                    : null;
            }
        }

        public MailSearchProgressDto? GetMailSearchProgressByCommandId(string commandId)
        {
            lock (_lock)
            {
                var progress = _mailSearchProgress.Values.LastOrDefault(item => item.CommandId == commandId);
                return progress is null ? null : CloneMailSearchProgress(progress);
            }
        }

        public void ApplyMailSearchSliceResult(MailSearchSliceResultDto batch)
        {
            lock (_lock)
            {
                var target = GetOrCreateMailSearchResultList(batch.SearchId);

                if (batch.Reset)
                {
                    target = new List<MailItemDto>();
                    if (!string.IsNullOrWhiteSpace(batch.SearchId))
                        _mailSearchResultsBySearchId[batch.SearchId] = target;
                }

                foreach (var mail in batch.Mails)
                    UpsertMail(target, CloneMailMetadata(mail));

                _mailSearchResults = target.Select(CloneMailMetadata).ToList();

                if (_mailSearchProgress.TryGetValue(batch.SearchId, out var progress))
                {
                    progress.Status = batch.IsFinal ? "completed" : "running";
                    progress.Phase = batch.IsFinal ? "completed" : "folder";
                    progress.ResultCount = target.Count;
                    progress.Message = batch.Message;
                    progress.Timestamp = DateTime.Now;
                }
            }
        }

        public void ApplyFolderMailsSliceResult(FolderMailsSliceResultDto batch)
        {
            lock (_lock)
            {
                var target = GetOrCreateFolderMailResultList(batch.FolderMailsId);
                if (batch.Reset)
                {
                    target = new List<MailItemDto>();
                    if (!string.IsNullOrWhiteSpace(batch.FolderMailsId))
                        _folderMailResultsByRequestId[batch.FolderMailsId] = target;
                }

                foreach (var mail in batch.Mails)
                    UpsertMail(target, CloneMailMetadata(mail));

                target.Sort((left, right) => right.ReceivedTime.CompareTo(left.ReceivedTime));
                _folderMailResults = target.Select(CloneMailMetadata).ToList();

                if (_mailSearchProgress.TryGetValue(batch.FolderMailsId, out var progress))
                {
                    progress.Status = batch.IsFinal ? "completed" : "running";
                    progress.Phase = batch.IsFinal ? "completed" : "folder";
                    progress.ResultCount = target.Count;
                    progress.Message = batch.Message;
                    progress.Timestamp = DateTime.Now;
                }
            }
        }

        public List<MailItemDto> GetMailSearchResults()
        {
            lock (_lock) { return _mailSearchResults.Select(CloneMail).ToList(); }
        }

        public List<MailItemDto> GetMailSearchResults(string searchId)
        {
            lock (_lock)
            {
                return !string.IsNullOrWhiteSpace(searchId) && _mailSearchResultsBySearchId.TryGetValue(searchId, out var results)
                    ? results.Select(CloneMail).ToList()
                    : _mailSearchResults.Select(CloneMail).ToList();
            }
        }

        public List<MailItemDto> GetFolderMailResults()
        {
            lock (_lock) { return _folderMailResults.Select(CloneMail).ToList(); }
        }

        public List<MailItemDto> GetFolderMailResults(string folderMailsId)
        {
            lock (_lock)
            {
                return !string.IsNullOrWhiteSpace(folderMailsId) && _folderMailResultsByRequestId.TryGetValue(folderMailsId, out var results)
                    ? results.Select(CloneMail).ToList()
                    : _folderMailResults.Select(CloneMail).ToList();
            }
        }

        public int GetMailSearchResultCount(string searchId)
        {
            lock (_lock)
            {
                return !string.IsNullOrWhiteSpace(searchId) && _mailSearchResultsBySearchId.TryGetValue(searchId, out var results)
                    ? results.Count
                    : _mailSearchResults.Count;
            }
        }

        public int GetFolderMailResultCount(string folderMailsId)
        {
            lock (_lock)
            {
                return !string.IsNullOrWhiteSpace(folderMailsId) && _folderMailResultsByRequestId.TryGetValue(folderMailsId, out var results)
                    ? results.Count
                    : _folderMailResults.Count;
            }
        }

        public FolderSnapshotDto GetFolderSnapshot()
        {
            lock (_lock)
            {
                return new FolderSnapshotDto
                {
                    Stores = CloneStores(_stores),
                    Folders = CloneFolders(_folders),
                };
            }
        }

        public void BeginFolderSync(bool reset = true)
        {
            if (!reset) return;
            lock (_lock)
            {
                _stores = new List<OutlookStoreDto>();
                _folders = new List<FolderDto>();
                _folderCacheUpdatedAt = DateTime.MinValue;
            }
        }

        public FolderSnapshotDto ApplyFolderBatch(FolderSyncBatchDto batch)
        {
            lock (_lock)
            {
                if (batch.Reset)
                {
                    _stores = new List<OutlookStoreDto>();
                    _folders = new List<FolderDto>();
                    _folderCacheUpdatedAt = DateTime.MinValue;
                }

                var rejectedStoreIds = batch.Stores
                    .Where(IsSystemStore)
                    .Select(store => store.StoreId)
                    .Where(storeId => !string.IsNullOrWhiteSpace(storeId))
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                if (rejectedStoreIds.Count > 0)
                {
                    _stores.RemoveAll(store => rejectedStoreIds.Contains(store.StoreId));
                    _folders.RemoveAll(folder => rejectedStoreIds.Contains(folder.StoreId));
                }

                foreach (var store in batch.Stores)
                {
                    if (rejectedStoreIds.Contains(store.StoreId)) continue;
                    UpsertStore(_stores, CloneStore(store));
                }

                foreach (var item in batch.Folders)
                {
                    if (rejectedStoreIds.Contains(item.StoreId)) continue;
                    if (IsRejectedFolder(item)) continue;
                    var folder = CloneFolder(item);
                    UpsertFolder(_folders, folder);
                }

                if (batch.Stores.Count > 0 || batch.Folders.Count > 0)
                    _folderCacheUpdatedAt = DateTime.Now;

                return new FolderSnapshotDto
                {
                    Stores = CloneStores(_stores),
                    Folders = CloneFolders(_folders),
                };
            }
        }

        public int CountFolders()
        {
            lock (_lock) { return CountFolders(_folders); }
        }

        public int CountStoreRoots()
        {
            lock (_lock) { return _folders.Count(folder => folder.IsStoreRoot); }
        }

        public bool IsFolderCacheStale(TimeSpan maxAge)
        {
            lock (_lock)
            {
                return _folderCacheUpdatedAt == DateTime.MinValue
                    || DateTime.Now - _folderCacheUpdatedAt > maxAge;
            }
        }

        public List<FolderDto> GetPendingFolderDiscoveryTargets()
        {
            lock (_lock)
            {
                return _folders
                    .Where(folder => folder.HasChildren && !folder.ChildrenLoaded)
                    .OrderBy(folder => folder.IsStoreRoot ? 0 : 1)
                    .ThenBy(folder => folder.StoreId)
                    .ThenBy(folder => folder.FolderPath)
                    .Select(CloneFolder)
                    .ToList();
            }
        }

        public bool IsFolderChildrenLoaded(string storeId, string parentEntryId, string parentFolderPath)
        {
            lock (_lock)
            {
                return _folders.Any(folder =>
                    !string.IsNullOrWhiteSpace(folder.StoreId)
                    && string.Equals(folder.StoreId, storeId, StringComparison.OrdinalIgnoreCase)
                    && (
                        (!string.IsNullOrWhiteSpace(parentEntryId) && string.Equals(folder.EntryId, parentEntryId, StringComparison.OrdinalIgnoreCase))
                        || (!string.IsNullOrWhiteSpace(parentFolderPath) && string.Equals(folder.FolderPath, parentFolderPath, StringComparison.OrdinalIgnoreCase))
                    )
                    && folder.ChildrenLoaded);
            }
        }

        public void SetRules(List<OutlookRuleDto> rules)
        {
            lock (_lock) { _rules = new List<OutlookRuleDto>(rules); }
        }

        public List<OutlookRuleDto> GetRules()
        {
            lock (_lock) { return new List<OutlookRuleDto>(_rules); }
        }

        public void SetCategories(List<OutlookCategoryDto> categories)
        {
            lock (_lock) { _categories = new List<OutlookCategoryDto>(categories); }
        }

        public List<OutlookCategoryDto> GetCategories()
        {
            lock (_lock) { return new List<OutlookCategoryDto>(_categories); }
        }

        public void SetCalendarEvents(List<CalendarEventDto> events)
        {
            lock (_lock) { _calendarEvents = new List<CalendarEventDto>(events); }
        }

        public List<CalendarEventDto> GetCalendarEvents()
        {
            lock (_lock) { return new List<CalendarEventDto>(_calendarEvents); }
        }

        private List<MailItemDto> GetOrCreateMailSearchResultList(string searchId)
        {
            if (string.IsNullOrWhiteSpace(searchId))
                return _mailSearchResults;

            if (!_mailSearchResultsBySearchId.TryGetValue(searchId, out var results))
            {
                results = new List<MailItemDto>();
                _mailSearchResultsBySearchId[searchId] = results;
            }

            return results;
        }

        private List<MailItemDto> GetOrCreateFolderMailResultList(string folderMailsId)
        {
            if (string.IsNullOrWhiteSpace(folderMailsId))
                return _folderMailResults;

            if (!_folderMailResultsByRequestId.TryGetValue(folderMailsId, out var results))
            {
                results = new List<MailItemDto>();
                _folderMailResultsByRequestId[folderMailsId] = results;
            }

            return results;
        }

        private static void UpsertStore(List<OutlookStoreDto> stores, OutlookStoreDto next)
        {
            var index = stores.FindIndex(store => store.StoreId == next.StoreId);
            if (index < 0) stores.Add(next);
            else stores[index] = next;
        }

        private static void UpsertFolder(List<FolderDto> folders, FolderDto next)
        {
            var index = folders.FindIndex(folder => folder.FolderPath == next.FolderPath);
            if (index < 0) folders.Add(next);
            else folders[index] = next;
        }

        private static void UpsertMail(List<MailItemDto> mails, MailItemDto next)
        {
            var index = mails.FindIndex(mail => mail.Id == next.Id);
            if (index < 0) mails.Add(next);
            else mails[index] = next;
        }

        private static void UpsertKnownMail(List<MailItemDto> mails, MailItemDto next)
        {
            var index = mails.FindIndex(mail => mail.Id == next.Id);
            if (index < 0) return;
            var merged = CloneMail(next);
            if (string.IsNullOrEmpty(merged.Body) && !string.IsNullOrEmpty(mails[index].Body))
                merged.Body = mails[index].Body;
            if (string.IsNullOrEmpty(merged.BodyHtml) && !string.IsNullOrEmpty(mails[index].BodyHtml))
                merged.BodyHtml = mails[index].BodyHtml;
            mails[index] = merged;
        }

        private static void UpdateMailBody(List<MailItemDto> mails, MailBodyDto body)
        {
            var mail = mails.FirstOrDefault(item => item.Id == body.MailId);
            if (mail is null) return;
            mail.Body = body.Body;
            mail.BodyHtml = body.BodyHtml;
        }

        private static int CountFolders(List<FolderDto> folders)
        {
            return folders.Count;
        }

        public static bool IsSearchableMailFolder(FolderDto folder)
        {
            return !folder.IsStoreRoot
                && folder.DefaultItemType == 0
                && !folder.IsHidden
                && !folder.IsSystem
                && IsAllowedMailFolderType(folder.FolderType);
        }

        private static bool IsRejectedFolder(FolderDto folder)
        {
            return folder.IsHidden
                || folder.IsSystem
                || !IsOperableFolder(folder);
        }

        private static bool IsOperableFolder(FolderDto folder)
        {
            return folder.IsStoreRoot || IsSearchableMailFolder(folder);
        }

        private static bool IsAllowedMailFolderType(OutlookFolderType folderType)
        {
            if (BlockedFolderTypes.Contains(folderType)) return false;

            return folderType is OutlookFolderType.Mail
                or OutlookFolderType.Inbox
                or OutlookFolderType.Sent
                or OutlookFolderType.Drafts
                or OutlookFolderType.Deleted
                or OutlookFolderType.Junk
                or OutlookFolderType.Archive
                or OutlookFolderType.Outbox;
        }

        private static bool IsSystemStore(OutlookStoreDto store)
        {
            return IsSystemNameOrPath(store.DisplayName) || IsSystemNameOrPath(store.RootFolderPath);
        }

        private static bool IsSystemNameOrPath(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return false;

            var segments = value
                .Split('\\', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

            return segments.Any(segment =>
                SystemFolderNames.Contains(segment)
                || SystemFolderNameFragments.Any(fragment => segment.Contains(fragment, StringComparison.OrdinalIgnoreCase)));
        }

        private static List<OutlookStoreDto> CloneStores(List<OutlookStoreDto> stores)
        {
            return stores.Select(CloneStore).ToList();
        }

        private static OutlookStoreDto CloneStore(OutlookStoreDto store)
        {
            return new OutlookStoreDto
            {
                StoreId = store.StoreId,
                DisplayName = store.DisplayName,
                StoreKind = store.StoreKind,
                StoreFilePath = store.StoreFilePath,
                RootFolderPath = store.RootFolderPath,
            };
        }

        private static List<FolderDto> CloneFolders(List<FolderDto> folders)
        {
            return folders.Select(CloneFolder).ToList();
        }

        private static MailItemDto CloneMail(MailItemDto mail)
        {
            return new MailItemDto
            {
                Id = mail.Id,
                Subject = mail.Subject,
                Sender = CloneRecipient(mail.Sender),
                ToRecipients = CloneRecipients(mail.ToRecipients),
                CcRecipients = CloneRecipients(mail.CcRecipients),
                BccRecipients = CloneRecipients(mail.BccRecipients),
                ReceivedTime = mail.ReceivedTime,
                Body = mail.Body,
                BodyHtml = mail.BodyHtml,
                FolderPath = mail.FolderPath,
                ConversationId = mail.ConversationId,
                ConversationTopic = mail.ConversationTopic,
                ConversationIndex = mail.ConversationIndex,
                Categories = mail.Categories,
                IsRead = mail.IsRead,
                IsMarkedAsTask = mail.IsMarkedAsTask,
                AttachmentCount = mail.AttachmentCount,
                AttachmentNames = mail.AttachmentNames,
                FlagRequest = mail.FlagRequest,
                FlagInterval = mail.FlagInterval,
                TaskStartDate = mail.TaskStartDate,
                TaskDueDate = mail.TaskDueDate,
                TaskCompletedDate = mail.TaskCompletedDate,
                Importance = mail.Importance,
                Sensitivity = mail.Sensitivity,
            };
        }

        private static List<OutlookRecipientDto> CloneRecipients(List<OutlookRecipientDto> recipients)
        {
            return recipients.Select(CloneRecipient).ToList();
        }

        private static OutlookRecipientDto CloneRecipient(OutlookRecipientDto recipient)
        {
            return new OutlookRecipientDto
            {
                RecipientKind = recipient.RecipientKind,
                DisplayName = recipient.DisplayName,
                SmtpAddress = recipient.SmtpAddress,
                RawAddress = recipient.RawAddress,
                AddressType = recipient.AddressType,
                EntryUserType = recipient.EntryUserType,
                IsGroup = recipient.IsGroup,
                IsResolved = recipient.IsResolved,
                Members = CloneRecipients(recipient.Members),
            };
        }

        private static MailItemDto CloneMailMetadata(MailItemDto mail)
        {
            var clone = CloneMail(mail);
            clone.Body = string.Empty;
            clone.BodyHtml = string.Empty;
            return clone;
        }

        private static FolderDto CloneFolder(FolderDto folder)
        {
            return new FolderDto
            {
                Name = folder.Name,
                EntryId = folder.EntryId,
                FolderPath = folder.FolderPath,
                ParentEntryId = folder.ParentEntryId,
                ParentFolderPath = folder.ParentFolderPath,
                ItemCount = folder.ItemCount,
                StoreId = folder.StoreId,
                IsStoreRoot = folder.IsStoreRoot,
                FolderType = folder.FolderType,
                DefaultItemType = folder.DefaultItemType,
                IsHidden = folder.IsHidden,
                IsSystem = folder.IsSystem,
                HasChildren = folder.HasChildren,
                ChildrenLoaded = folder.ChildrenLoaded,
                DiscoveryState = folder.DiscoveryState,
            };
        }

        private static MailSearchProgressDto CloneMailSearchProgress(MailSearchProgressDto progress)
        {
            return new MailSearchProgressDto
            {
                SearchId = progress.SearchId,
                CommandId = progress.CommandId,
                Status = progress.Status,
                Phase = progress.Phase,
                ProcessedStores = progress.ProcessedStores,
                TotalStores = progress.TotalStores,
                ProcessedFolders = progress.ProcessedFolders,
                TotalFolders = progress.TotalFolders,
                ResultCount = progress.ResultCount,
                CurrentStoreId = progress.CurrentStoreId,
                CurrentFolderPath = progress.CurrentFolderPath,
                Message = progress.Message,
                Timestamp = progress.Timestamp,
            };
        }

        private static SearchMailsRequest CloneSearchMailsRequest(SearchMailsRequest request)
        {
            return new SearchMailsRequest
            {
                SearchId = request.SearchId,
                StoreId = request.StoreId,
                ScopeFolderPaths = new List<string>(request.ScopeFolderPaths),
                AllowGlobalScope = request.AllowGlobalScope,
                IncludeSubFolders = request.IncludeSubFolders,
                Keyword = request.Keyword,
                TextFields = new List<string>(request.TextFields),
                CategoryNames = new List<string>(request.CategoryNames),
                HasAttachments = request.HasAttachments,
                FlagState = request.FlagState,
                ReadState = request.ReadState,
                ReceivedFrom = request.ReceivedFrom,
                ReceivedTo = request.ReceivedTo,
            };
        }

        private static MailAttachmentsDto CloneMailAttachments(MailAttachmentsDto attachments)
        {
            return new MailAttachmentsDto
            {
                MailId = attachments.MailId,
                FolderPath = attachments.FolderPath,
                Attachments = attachments.Attachments.Select(CloneMailAttachment).ToList(),
            };
        }

        private static MailConversationDto CloneMailConversation(MailConversationDto conversation)
        {
            return new MailConversationDto
            {
                MailId = conversation.MailId,
                FolderPath = conversation.FolderPath,
                ConversationId = conversation.ConversationId,
                ConversationTopic = conversation.ConversationTopic,
                Mails = conversation.Mails.Select(CloneMail).ToList(),
            };
        }

        private static MailAttachmentDto CloneMailAttachment(MailAttachmentDto attachment)
        {
            return new MailAttachmentDto
            {
                MailId = attachment.MailId,
                Id = attachment.Id,
                AttachmentId = attachment.AttachmentId,
                Index = attachment.Index,
                FileName = attachment.FileName,
                DisplayName = attachment.DisplayName,
                Name = attachment.Name,
                ContentType = attachment.ContentType,
                Size = attachment.Size,
                IsExported = attachment.IsExported,
                ExportedAttachmentId = attachment.ExportedAttachmentId,
                Path = attachment.Path,
                LocalPath = attachment.LocalPath,
                FullPath = attachment.FullPath,
                ExportedPath = attachment.ExportedPath,
            };
        }

        private static ExportedMailAttachmentDto CloneExportedAttachment(ExportedMailAttachmentDto attachment)
        {
            return new ExportedMailAttachmentDto
            {
                MailId = attachment.MailId,
                FolderPath = attachment.FolderPath,
                Id = attachment.Id,
                AttachmentId = attachment.AttachmentId,
                Index = attachment.Index,
                ExportedAttachmentId = attachment.ExportedAttachmentId,
                FileName = attachment.FileName,
                DisplayName = attachment.DisplayName,
                Name = attachment.Name,
                ContentType = attachment.ContentType,
                Size = attachment.Size,
                Path = attachment.Path,
                LocalPath = attachment.LocalPath,
                FullPath = attachment.FullPath,
                ExportedPath = attachment.ExportedPath,
                ExportedAt = attachment.ExportedAt,
            };
        }

        private static void NormalizeMailAttachments(MailAttachmentsDto attachments)
        {
            foreach (var attachment in attachments.Attachments)
            {
                if (string.IsNullOrWhiteSpace(attachment.MailId))
                    attachment.MailId = attachments.MailId;
                NormalizeMailAttachment(attachment);
            }
        }

        private static void NormalizeMailConversation(MailConversationDto conversation)
        {
            conversation.Mails ??= new List<MailItemDto>();
            foreach (var mail in conversation.Mails)
            {
                if (string.IsNullOrWhiteSpace(mail.ConversationId))
                    mail.ConversationId = conversation.ConversationId;
                if (string.IsNullOrWhiteSpace(mail.ConversationTopic))
                    mail.ConversationTopic = conversation.ConversationTopic;
            }
        }

        private static void NormalizeMailAttachment(MailAttachmentDto attachment)
        {
            if (string.IsNullOrWhiteSpace(attachment.AttachmentId))
                attachment.AttachmentId = FirstNonBlank(attachment.Id, attachment.Index > 0 ? attachment.Index.ToString() : string.Empty);
            if (string.IsNullOrWhiteSpace(attachment.Id))
                attachment.Id = attachment.AttachmentId;
            if (attachment.Index <= 0 && int.TryParse(attachment.AttachmentId, out var index))
                attachment.Index = index;

            if (string.IsNullOrWhiteSpace(attachment.Name))
                attachment.Name = FirstNonBlank(attachment.FileName, attachment.DisplayName);
            if (string.IsNullOrWhiteSpace(attachment.FileName))
                attachment.FileName = attachment.Name;
            if (string.IsNullOrWhiteSpace(attachment.DisplayName))
                attachment.DisplayName = attachment.Name;

            if (string.IsNullOrWhiteSpace(attachment.ExportedPath))
                attachment.ExportedPath = FirstNonBlank(attachment.LocalPath, attachment.FullPath, attachment.Path);
            if (string.IsNullOrWhiteSpace(attachment.LocalPath))
                attachment.LocalPath = attachment.ExportedPath;
            if (string.IsNullOrWhiteSpace(attachment.FullPath))
                attachment.FullPath = attachment.ExportedPath;
            if (string.IsNullOrWhiteSpace(attachment.Path))
                attachment.Path = attachment.ExportedPath;
        }

        private static void NormalizeExportedAttachment(ExportedMailAttachmentDto attachment)
        {
            if (string.IsNullOrWhiteSpace(attachment.AttachmentId))
                attachment.AttachmentId = FirstNonBlank(attachment.Id, attachment.Index > 0 ? attachment.Index.ToString() : string.Empty);
            if (string.IsNullOrWhiteSpace(attachment.Id))
                attachment.Id = attachment.AttachmentId;
            if (attachment.Index <= 0 && int.TryParse(attachment.AttachmentId, out var index))
                attachment.Index = index;
            if (string.IsNullOrWhiteSpace(attachment.ExportedAttachmentId))
                attachment.ExportedAttachmentId = Guid.NewGuid().ToString();

            if (string.IsNullOrWhiteSpace(attachment.Name))
                attachment.Name = FirstNonBlank(attachment.FileName, attachment.DisplayName);
            if (string.IsNullOrWhiteSpace(attachment.FileName))
                attachment.FileName = attachment.Name;
            if (string.IsNullOrWhiteSpace(attachment.DisplayName))
                attachment.DisplayName = attachment.Name;

            if (string.IsNullOrWhiteSpace(attachment.ExportedPath))
                attachment.ExportedPath = FirstNonBlank(attachment.LocalPath, attachment.FullPath, attachment.Path);
            if (string.IsNullOrWhiteSpace(attachment.LocalPath))
                attachment.LocalPath = attachment.ExportedPath;
            if (string.IsNullOrWhiteSpace(attachment.FullPath))
                attachment.FullPath = attachment.ExportedPath;
            if (string.IsNullOrWhiteSpace(attachment.Path))
                attachment.Path = attachment.ExportedPath;
        }

        private static string FirstNonBlank(params string[] values)
        {
            return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? string.Empty;
        }
    }

    public class ChatStore
    {
        private readonly List<ChatMessageDto> _messages = new();
        private readonly object _lock = new();

        public void Add(ChatMessageDto msg)
        {
            lock (_lock) { _messages.Add(msg); }
        }

        public List<ChatMessageDto> GetAll()
        {
            lock (_lock) { return new List<ChatMessageDto>(_messages); }
        }
    }

    public class CommandResultStore
    {
        private readonly object _lock = new();
        private readonly Dictionary<string, OutlookCommandStatusDto> _commands = new();
        private readonly Dictionary<string, PendingCommand> _requestCommands = new();
        private readonly Queue<string> _order = new();
        private const int MaxCommands = 500;

        public void RecordDispatched(PendingCommand command)
        {
            lock (_lock)
            {
                if (!_commands.ContainsKey(command.Id))
                    _order.Enqueue(command.Id);

                _commands[command.Id] = new OutlookCommandStatusDto
                {
                    CommandId = command.Id,
                    Type = command.Type,
                    Status = "pending",
                    DispatchTimestamp = DateTime.Now,
                };
                _requestCommands[command.Id] = command;

                TrimIfNeeded();
            }
        }

        public void RecordUnavailable(PendingCommand command)
        {
            lock (_lock)
            {
                if (!_commands.TryGetValue(command.Id, out var status))
                {
                    status = new OutlookCommandStatusDto
                    {
                        CommandId = command.Id,
                        Type = command.Type,
                        DispatchTimestamp = DateTime.Now,
                    };
                    _commands[command.Id] = status;
                    _order.Enqueue(command.Id);
                }
                _requestCommands[command.Id] = command;

                status.Status = "addin_unavailable";
                status.Success = false;
                status.Message = "No Outlook AddIn SignalR connection is available.";
                status.ResultTimestamp = DateTime.Now;
                TrimIfNeeded();
            }
        }

        public void RecordResult(OutlookCommandResult result)
        {
            lock (_lock)
            {
                if (!_commands.TryGetValue(result.CommandId, out var status))
                {
                    status = new OutlookCommandStatusDto
                    {
                        CommandId = result.CommandId,
                        DispatchTimestamp = result.Timestamp,
                    };
                    _commands[result.CommandId] = status;
                    _order.Enqueue(result.CommandId);
                }

                status.Status = result.Success ? "completed" : "failed";
                status.Success = result.Success;
                status.Message = result.Message;
                status.Payload = result.Payload;
                status.ResultTimestamp = result.Timestamp;
                TrimIfNeeded();
            }
        }

        public OutlookCommandStatusDto? Get(string commandId)
        {
            lock (_lock)
            {
                if (!_commands.TryGetValue(commandId, out var status))
                    return null;

                return Clone(status);
            }
        }

        public PendingCommand? GetRequestCommand(string requestId)
        {
            lock (_lock)
            {
                return _requestCommands.TryGetValue(requestId, out var command) ? command : null;
            }
        }

        public List<OutlookCommandStatusDto> GetRecent()
        {
            lock (_lock)
            {
                return _order
                    .Where(_commands.ContainsKey)
                    .Select(id => Clone(_commands[id]))
                    .Reverse()
                    .ToList();
            }
        }

        private void TrimIfNeeded()
        {
            while (_order.Count > MaxCommands)
            {
                var oldestId = _order.Dequeue();
                _commands.Remove(oldestId);
                _requestCommands.Remove(oldestId);
            }
        }

        private static OutlookCommandStatusDto Clone(OutlookCommandStatusDto status)
        {
            return new OutlookCommandStatusDto
            {
                CommandId = status.CommandId,
                Type = status.Type,
                Status = status.Status,
                Success = status.Success,
                Message = status.Message,
                Payload = status.Payload,
                DispatchTimestamp = status.DispatchTimestamp,
                ResultTimestamp = status.ResultTimestamp,
            };
        }
    }

    public class AddinStatusStore
    {
        private readonly object _lock = new();
        private readonly List<AddinLogEntry> _logs = new();
        private readonly HashSet<string> _signalRConnectionIds = new();
        private bool _mockBackendActive;
        private bool _connected;
        private DateTime? _lastPollTime;
        private DateTime? _lastPushTime;
        private string _lastCommand = string.Empty;

        public void RecordPush(string pushType, int count)
        {
            lock (_lock)
            {
                _connected = true;
                _lastPushTime = DateTime.Now;
                AddLogInternal("info", $"Add-in pushed {pushType}: {count} items");
            }
        }

        public void RecordSignalRConnected(string connectionId, string clientName)
        {
            lock (_lock)
            {
                _signalRConnectionIds.Add(connectionId);
                _connected = true;
                _lastPollTime = DateTime.Now;
                AddLogInternal("info", $"Add-in SignalR connected: {clientName}");
            }
        }

        public void RecordSignalRDisconnected(string connectionId)
        {
            lock (_lock)
            {
                _signalRConnectionIds.Remove(connectionId);
                _connected = _signalRConnectionIds.Count > 0;
                AddLogInternal("warn", $"Add-in SignalR disconnected: {connectionId}");
            }
        }

        public void RecordSignalRDispatch(string commandType)
        {
            lock (_lock)
            {
                _connected = true;
                _lastCommand = commandType;
                AddLogInternal("info", $"Add-in SignalR command dispatched: {commandType}");
            }
        }

        public void RecordMockDispatch(string commandType)
        {
            lock (_lock)
            {
                _mockBackendActive = true;
                _connected = true;
                _lastCommand = commandType;
                _lastPushTime = DateTime.Now;
                AddLogInternal("info", $"Mock Outlook command handled: {commandType}");
            }
        }

        public void RecordMockBackendReady()
        {
            lock (_lock)
            {
                _mockBackendActive = true;
                _connected = true;
                _lastPollTime = DateTime.Now;
                AddLogInternal("info", "Mock Outlook backend ready");
            }
        }

        public void AddLog(string level, string message)
        {
            lock (_lock) { AddLogInternal(level, message); }
        }

        private void AddLogInternal(string level, string message)
        {
            _logs.Add(new AddinLogEntry { Level = level, Message = message, Timestamp = DateTime.Now });
            if (_logs.Count > 500) _logs.RemoveAt(0);
        }

        public AddinStatusDto GetStatus()
        {
            lock (_lock)
            {
                // Mock backend 或 SignalR 連線存在時，都視為可處理 Outlook command。
                if (_mockBackendActive)
                    _connected = true;
                else if (_signalRConnectionIds.Count > 0)
                    _connected = true;
                else if (_lastPollTime.HasValue && (DateTime.Now - _lastPollTime.Value).TotalSeconds > 90)
                    _connected = false;

                return new AddinStatusDto
                {
                    Connected = _connected,
                    LastPollTime = _lastPollTime,
                    LastPushTime = _lastPushTime,
                    LastCommand = _lastCommand
                };
            }
        }

        public bool HasSignalRConnection()
        {
            lock (_lock) { return _signalRConnectionIds.Count > 0; }
        }

        public List<AddinLogEntry> GetLogs()
        {
            lock (_lock) { return new List<AddinLogEntry>(_logs); }
        }
    }

    public class OutlookSignalRCommandDispatcher
    {
        private readonly AddinStatusStore _addinStatus;
        private readonly Microsoft.AspNetCore.SignalR.IHubContext<Hubs.OutlookAddinHub> _hub;

        public OutlookSignalRCommandDispatcher(
            AddinStatusStore addinStatus,
            Microsoft.AspNetCore.SignalR.IHubContext<Hubs.OutlookAddinHub> hub)
        {
            _addinStatus = addinStatus;
            _hub = hub;
        }

        public async Task<bool> DispatchAsync(PendingCommand command, CancellationToken cancellationToken = default)
        {
            if (!_addinStatus.HasSignalRConnection())
                return false;

            await _hub.Clients
                .Group(Hubs.OutlookAddinHub.AddinGroupName)
                .SendAsync("OutlookCommand", command, cancellationToken);
            _addinStatus.RecordSignalRDispatch(command.Type);
            return true;
        }
    }
}
