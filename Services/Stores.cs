using Microsoft.AspNetCore.SignalR;
using SmartOffice.Hub.Models;

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

        public void ApplyMailSearchBatch(MailSearchBatchDto batch)
        {
            lock (_lock)
            {
                if (batch.Reset) _mailSearchResults = new List<MailItemDto>();
                foreach (var mail in batch.Mails)
                    UpsertMail(_mailSearchResults, CloneMail(mail));
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
                FolderPath = folder.FolderPath,
                ParentFolderPath = folder.ParentFolderPath,
                ItemCount = folder.ItemCount,
                StoreId = folder.StoreId,
                IsStoreRoot = folder.IsStoreRoot,
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

        private static MailAttachmentDto CloneMailAttachment(MailAttachmentDto attachment)
        {
            return new MailAttachmentDto
            {
                MailId = attachment.MailId,
                AttachmentId = attachment.AttachmentId,
                Name = attachment.Name,
                ContentType = attachment.ContentType,
                Size = attachment.Size,
                IsExported = attachment.IsExported,
                ExportedAttachmentId = attachment.ExportedAttachmentId,
                ExportedPath = attachment.ExportedPath,
            };
        }

        private static ExportedMailAttachmentDto CloneExportedAttachment(ExportedMailAttachmentDto attachment)
        {
            return new ExportedMailAttachmentDto
            {
                MailId = attachment.MailId,
                FolderPath = attachment.FolderPath,
                AttachmentId = attachment.AttachmentId,
                ExportedAttachmentId = attachment.ExportedAttachmentId,
                Name = attachment.Name,
                ContentType = attachment.ContentType,
                Size = attachment.Size,
                ExportedPath = attachment.ExportedPath,
                ExportedAt = attachment.ExportedAt,
            };
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
