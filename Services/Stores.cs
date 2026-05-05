using Microsoft.AspNetCore.SignalR;
using SmartOffice.Hub.Models;
using System.Text.RegularExpressions;

namespace SmartOffice.Hub.Services
{
    public class MailStore
    {
        private readonly object _lock = new();
        private List<MailItemDto> _mails = new();
        private List<FolderDto> _folders = new();
        private List<OutlookStoreDto> _stores = new();
        private List<OutlookRuleDto> _rules = new();
        private List<OutlookCategoryDto> _categories = new();
        private List<CalendarEventDto> _calendarEvents = new();
        private List<MailItemDto> _mailSearchResults = new();
        private readonly Dictionary<string, MailSearchProgressDto> _mailSearchProgress = new();
        private readonly Dictionary<string, SearchMailsRequest> _mailSearchRequests = new();
        private readonly Dictionary<string, MailAttachmentsDto> _attachments = new();
        private readonly Dictionary<string, ExportedMailAttachmentDto> _exportedAttachments = new();

        public void SetMails(List<MailItemDto> mails)
        {
            lock (_lock) { _mails = new List<MailItemDto>(mails); }
        }

        public void UpsertMail(MailItemDto mail)
        {
            lock (_lock)
            {
                var index = _mails.FindIndex(item => item.Id == mail.Id);
                if (index < 0) return;
                if (string.IsNullOrEmpty(mail.Body) && !string.IsNullOrEmpty(_mails[index].Body))
                    mail.Body = _mails[index].Body;
                if (string.IsNullOrEmpty(mail.BodyHtml) && !string.IsNullOrEmpty(_mails[index].BodyHtml))
                    mail.BodyHtml = _mails[index].BodyHtml;
                _mails[index] = mail;
            }
        }

        public void UpdateMailBody(MailBodyDto body)
        {
            lock (_lock)
            {
                var mail = _mails.FirstOrDefault(item => item.Id == body.MailId);
                if (mail is null) return;
                mail.Body = body.Body;
                mail.BodyHtml = body.BodyHtml;
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
            lock (_lock) { return new List<MailItemDto>(_mails); }
        }

        public void BeginMailSearch(bool reset = true)
        {
            if (!reset) return;
            lock (_lock) { _mailSearchResults = new List<MailItemDto>(); }
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
                if (batch.Reset) _mailSearchResults = new List<MailItemDto>();
                var mails = _mailSearchRequests.TryGetValue(batch.SearchId, out var request)
                    ? batch.Mails.Where(mail => MatchesSearchRequest(mail, request))
                    : batch.Mails;
                foreach (var mail in mails)
                    UpsertMail(_mailSearchResults, CloneMail(mail));

                if (_mailSearchProgress.TryGetValue(batch.SearchId, out var progress))
                {
                    progress.Status = batch.IsFinal ? "completed" : "running";
                    progress.Phase = batch.IsFinal ? "completed" : "folder";
                    progress.ResultCount = _mailSearchResults.Count;
                    progress.Message = batch.Message;
                    progress.Timestamp = DateTime.Now;
                }
            }
        }

        public List<MailItemDto> GetMailSearchResults()
        {
            lock (_lock) { return _mailSearchResults.Select(CloneMail).ToList(); }
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
                }

                foreach (var store in batch.Stores)
                    UpsertStore(_stores, CloneStore(store));

                foreach (var item in batch.Folders)
                {
                    var folder = CloneFolder(item);
                    UpsertFolder(_folders, folder);
                }

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

        private static int CountFolders(List<FolderDto> folders)
        {
            return folders.Count;
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
                SenderName = mail.SenderName,
                SenderEmail = mail.SenderEmail,
                ReceivedTime = mail.ReceivedTime,
                Body = mail.Body,
                BodyHtml = mail.BodyHtml,
                FolderPath = mail.FolderPath,
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
                IncludeSubFolders = request.IncludeSubFolders,
                Keyword = request.Keyword,
                MatchMode = request.MatchMode,
                Fields = new List<string>(request.Fields),
                ReceivedFrom = request.ReceivedFrom,
                ReceivedTo = request.ReceivedTo,
                MaxCount = request.MaxCount,
            };
        }

        private static bool MatchesSearchRequest(MailItemDto mail, SearchMailsRequest request)
        {
            var keyword = request.Keyword.Trim();
            if (string.IsNullOrWhiteSpace(keyword)) return true;
            var haystack = SearchHaystack(mail, request.Fields);
            if (string.Equals(request.MatchMode, "exact", StringComparison.OrdinalIgnoreCase))
                return haystack.Any(value => string.Equals(value, keyword, StringComparison.OrdinalIgnoreCase));
            if (string.Equals(request.MatchMode, "regex", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    return haystack.Any(value => Regex.IsMatch(value, keyword, RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(100)));
                }
                catch
                {
                    return false;
                }
            }
            if (string.Equals(request.MatchMode, "fuzzy", StringComparison.OrdinalIgnoreCase))
                return haystack.Any(value => FuzzyMatches(value, keyword));
            return haystack.Any(value => value.Contains(keyword, StringComparison.OrdinalIgnoreCase));
        }

        private static List<string> SearchHaystack(MailItemDto mail, List<string> fields)
        {
            var selected = fields.Count == 0 ? new List<string> { "subject" } : fields;
            var values = new List<string>();
            if (HasSearchField(selected, "subject")) values.Add(mail.Subject);
            if (HasSearchField(selected, "sender"))
            {
                values.Add(mail.SenderName);
                values.Add(mail.SenderEmail);
            }
            if (HasSearchField(selected, "categories")) values.Add(mail.Categories);
            if (HasSearchField(selected, "body"))
            {
                values.Add(mail.Body);
                values.Add(mail.BodyHtml);
            }
            return values;
        }

        private static bool HasSearchField(List<string> fields, string field)
        {
            return fields.Any(item => string.Equals(item, field, StringComparison.OrdinalIgnoreCase));
        }

        private static bool FuzzyMatches(string value, string keyword)
        {
            var normalizedValue = NormalizeFuzzyText(value);
            var normalizedKeyword = NormalizeFuzzyText(keyword);
            if (string.IsNullOrWhiteSpace(normalizedValue) || string.IsNullOrWhiteSpace(normalizedKeyword)) return false;
            if (normalizedValue.Contains(normalizedKeyword, StringComparison.OrdinalIgnoreCase)) return true;
            if (IsSubsequence(normalizedKeyword, normalizedValue)) return true;

            var keywordTokens = SplitFuzzyTokens(keyword);
            if (keywordTokens.Count == 0) return false;
            var valueTokens = SplitFuzzyTokens(value);
            return keywordTokens.All(term => valueTokens.Any(token => WithinFuzzyDistance(token, term)));
        }

        private static string NormalizeFuzzyText(string value)
        {
            return new string(value
                .Where(char.IsLetterOrDigit)
                .Select(char.ToLowerInvariant)
                .ToArray());
        }

        private static List<string> SplitFuzzyTokens(string value)
        {
            return Regex.Split(value.ToLowerInvariant(), @"[^\p{L}\p{Nd}]+")
                .Where(token => !string.IsNullOrWhiteSpace(token))
                .ToList();
        }

        private static bool WithinFuzzyDistance(string value, string keyword)
        {
            if (value.Contains(keyword, StringComparison.OrdinalIgnoreCase)) return true;
            var threshold = keyword.Length <= 4 ? 1 : Math.Max(1, keyword.Length / 4);
            return LevenshteinDistance(value, keyword) <= threshold;
        }

        private static bool IsSubsequence(string needle, string haystack)
        {
            var index = 0;
            foreach (var item in haystack)
            {
                if (index < needle.Length && item == needle[index]) index++;
                if (index == needle.Length) return true;
            }
            return false;
        }

        private static int LevenshteinDistance(string left, string right)
        {
            if (left.Length == 0) return right.Length;
            if (right.Length == 0) return left.Length;

            var previous = Enumerable.Range(0, right.Length + 1).ToArray();
            var current = new int[right.Length + 1];
            for (var i = 1; i <= left.Length; i++)
            {
                current[0] = i;
                for (var j = 1; j <= right.Length; j++)
                {
                    var cost = left[i - 1] == right[j - 1] ? 0 : 1;
                    current[j] = Math.Min(
                        Math.Min(current[j - 1] + 1, previous[j] + 1),
                        previous[j - 1] + cost);
                }
                (previous, current) = (current, previous);
            }
            return previous[right.Length];
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
