using Microsoft.AspNetCore.SignalR;
using SmartOffice.Hub.Models;

namespace SmartOffice.Hub.Services
{
    public partial class MailStore
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

        public MailItemDto? FindCachedMail(string mailId)
        {
            if (string.IsNullOrWhiteSpace(mailId)) return null;
            lock (_lock)
            {
                return _mails.Concat(_mailSearchResults)
                    .Concat(_folderMailResults)
                    .Concat(_mailSearchResultsBySearchId.Values.SelectMany(items => items))
                    .Concat(_folderMailResultsByRequestId.Values.SelectMany(items => items))
                    .FirstOrDefault(mail => string.Equals(mail.Id, mailId, StringComparison.OrdinalIgnoreCase)) is { } found
                        ? CloneMail(found)
                        : null;
            }
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

}
