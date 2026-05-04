using Microsoft.AspNetCore.SignalR;
using SmartOffice.Hub.Models;
using SmartOffice.Hub.Services;

namespace SmartOffice.Hub.Hubs
{
    /// <summary>
    /// Outlook AddIn 正式 SignalR channel。
    /// Web UI / AI 仍透過 HTTP API 呼叫 Hub；Hub 透過這個 channel
    /// 即時 dispatch command 給已連線的 Outlook AddIn。
    /// </summary>
    public class OutlookAddinHub : Microsoft.AspNetCore.SignalR.Hub
    {
        public const string AddinGroupName = "outlook-addins";

        private readonly MailStore _mailStore;
        private readonly ChatStore _chatStore;
        private readonly CommandResultStore _commandResults;
        private readonly AddinStatusStore _addinStatus;
        private readonly IHubContext<NotificationHub> _notifications;

        public OutlookAddinHub(
            MailStore mailStore,
            ChatStore chatStore,
            CommandResultStore commandResults,
            AddinStatusStore addinStatus,
            IHubContext<NotificationHub> notifications)
        {
            _mailStore = mailStore;
            _chatStore = chatStore;
            _commandResults = commandResults;
            _addinStatus = addinStatus;
            _notifications = notifications;
        }

        public async Task RegisterOutlookAddin(OutlookAddinClientInfo info)
        {
            await Groups.AddToGroupAsync(Context.ConnectionId, AddinGroupName);
            var clientName = string.IsNullOrWhiteSpace(info.ClientName) ? "Outlook AddIn" : info.ClientName;
            _addinStatus.RecordSignalRConnected(Context.ConnectionId, clientName);
            await BroadcastStatusAndLogsAsync();
        }

        public async Task BeginFolderSync(FolderSyncBeginDto info)
        {
            _mailStore.BeginFolderSync(info.Reset);
            _addinStatus.AddLog("info", $"Folder sync started: {info.SyncId}");
            await _notifications.Clients.All.SendAsync("FolderSyncStarted", info);
            await BroadcastStatusAndLogsAsync();
        }

        public async Task PushFolderBatch(FolderSyncBatchDto batch)
        {
            _mailStore.ApplyFolderBatch(batch);
            _addinStatus.RecordPush("folder batch", batch.Stores.Count + batch.Folders.Count);
            await _notifications.Clients.All.SendAsync("FoldersPatched", batch);

            if (batch.IsFinal)
            {
                var complete = new FolderSyncCompleteDto
                {
                    SyncId = batch.SyncId,
                    TotalCount = _mailStore.CountFolders(),
                    Message = "Folder sync completed by final batch",
                };
                await _notifications.Clients.All.SendAsync("FolderSyncCompleted", complete);
            }

            await BroadcastStatusAndLogsAsync();
        }

        public async Task CompleteFolderSync(FolderSyncCompleteDto info)
        {
            if (info.TotalCount <= 0)
                info.TotalCount = _mailStore.CountFolders();

            _addinStatus.AddLog(info.Success ? "info" : "warn", $"Folder sync completed: {info.TotalCount} folders. {info.Message}");
            await _notifications.Clients.All.SendAsync("FolderSyncCompleted", info);
            await BroadcastStatusAndLogsAsync();
        }

        public async Task PushMails(List<MailItemDto> mails)
        {
            _mailStore.SetMails(mails);
            _addinStatus.RecordPush("mails", mails.Count);
            await _notifications.Clients.All.SendAsync("MailsUpdated", mails);
            await BroadcastStatusAndLogsAsync();
        }

        public async Task PushMail(MailItemDto mail)
        {
            _mailStore.UpsertMail(mail);
            _addinStatus.RecordPush("mail", 1);
            await _notifications.Clients.All.SendAsync("MailUpdated", mail);
            await BroadcastStatusAndLogsAsync();
        }

        public async Task BeginMailSearch(MailSearchBatchDto batch)
        {
            _mailStore.BeginMailSearch(batch.Reset);
            _addinStatus.AddLog("info", $"Mail search started: {batch.SearchId}");
            await _notifications.Clients.All.SendAsync("MailSearchStarted", batch);
            await BroadcastStatusAndLogsAsync();
        }

        public async Task PushMailSearchBatch(MailSearchBatchDto batch)
        {
            _mailStore.ApplyMailSearchBatch(batch);
            _addinStatus.RecordPush("mail search results", batch.Mails.Count);
            await _notifications.Clients.All.SendAsync("MailSearchPatched", batch);

            if (batch.IsFinal)
            {
                var complete = new MailSearchCompleteDto
                {
                    SearchId = batch.SearchId,
                    TotalCount = _mailStore.GetMailSearchResults().Count,
                    Message = string.IsNullOrWhiteSpace(batch.Message) ? "Mail search completed by final batch" : batch.Message,
                };
                await _notifications.Clients.All.SendAsync("MailSearchCompleted", complete);
            }

            await BroadcastStatusAndLogsAsync();
        }

        public async Task CompleteMailSearch(MailSearchCompleteDto info)
        {
            if (info.TotalCount <= 0)
                info.TotalCount = _mailStore.GetMailSearchResults().Count;

            _addinStatus.AddLog(info.Success ? "info" : "warn", $"Mail search completed: {info.TotalCount} mails. {info.Message}");
            await _notifications.Clients.All.SendAsync("MailSearchCompleted", info);
            await BroadcastStatusAndLogsAsync();
        }

        public async Task PushMailBody(MailBodyDto body)
        {
            _mailStore.UpdateMailBody(body);
            _addinStatus.RecordPush("mail body", 1);
            await _notifications.Clients.All.SendAsync("MailBodyUpdated", body);
            await BroadcastStatusAndLogsAsync();
        }

        public async Task PushMailAttachments(MailAttachmentsDto attachments)
        {
            _mailStore.SetMailAttachments(attachments);
            _addinStatus.RecordPush("mail attachments", attachments.Attachments.Count);
            await _notifications.Clients.All.SendAsync("MailAttachmentsUpdated", attachments);
            await BroadcastStatusAndLogsAsync();
        }

        public async Task PushExportedMailAttachment(ExportedMailAttachmentDto attachment)
        {
            _mailStore.UpsertExportedAttachment(attachment);
            _addinStatus.RecordPush("exported attachment", 1);
            await _notifications.Clients.All.SendAsync("MailAttachmentExported", attachment);
            await BroadcastStatusAndLogsAsync();
        }

        public async Task PushRules(List<OutlookRuleDto> rules)
        {
            _mailStore.SetRules(rules);
            _addinStatus.RecordPush("rules", rules.Count);
            await _notifications.Clients.All.SendAsync("RulesUpdated", rules);
            await BroadcastStatusAndLogsAsync();
        }

        public async Task PushCategories(List<OutlookCategoryDto> categories)
        {
            _mailStore.SetCategories(categories);
            _addinStatus.RecordPush("categories", categories.Count);
            await _notifications.Clients.All.SendAsync("CategoriesUpdated", categories);
            await BroadcastStatusAndLogsAsync();
        }

        public async Task PushCalendar(List<CalendarEventDto> events)
        {
            _mailStore.SetCalendarEvents(events);
            _addinStatus.RecordPush("calendar", events.Count);
            await _notifications.Clients.All.SendAsync("CalendarUpdated", events);
            await BroadcastStatusAndLogsAsync();
        }

        public async Task ReportAddinLog(AddinLogEntry entry)
        {
            _addinStatus.AddLog(entry.Level, entry.Message);
            await _notifications.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs());
        }

        public async Task SendChatMessage(ChatMessageDto message)
        {
            if (string.IsNullOrWhiteSpace(message.Source))
                message.Source = "outlook";

            message.Timestamp = DateTime.Now;
            _chatStore.Add(message);
            await _notifications.Clients.All.SendAsync("NewChatMessage", message);
        }

        public async Task ReportCommandResult(OutlookCommandResult result)
        {
            _commandResults.RecordResult(result);
            var level = result.Success ? "info" : "warn";
            _addinStatus.AddLog(level, $"Command result {result.CommandId}: {result.Message}");
            await _notifications.Clients.All.SendAsync("CommandResult", result);
            await BroadcastStatusAndLogsAsync();
        }

        public override async Task OnDisconnectedAsync(Exception? exception)
        {
            _addinStatus.RecordSignalRDisconnected(Context.ConnectionId);
            await BroadcastStatusAndLogsAsync();
            await base.OnDisconnectedAsync(exception);
        }

        private async Task BroadcastStatusAndLogsAsync()
        {
            await _notifications.Clients.All.SendAsync("AddinStatus", _addinStatus.GetStatus());
            await _notifications.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs());
        }
    }
}
