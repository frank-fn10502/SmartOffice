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
        private readonly AddinStatusStore _addinStatus;
        private readonly IHubContext<NotificationHub> _notifications;

        public OutlookAddinHub(
            MailStore mailStore,
            AddinStatusStore addinStatus,
            IHubContext<NotificationHub> notifications)
        {
            _mailStore = mailStore;
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

        public async Task PushFolders(List<FolderDto> folders)
        {
            _mailStore.SetFolders(folders);
            _addinStatus.RecordPush("folders", folders.Count);
            await _notifications.Clients.All.SendAsync("FoldersUpdated", folders);
            await BroadcastStatusAndLogsAsync();
        }

        public async Task PushMails(List<MailItemDto> mails)
        {
            _mailStore.SetMails(mails);
            _addinStatus.RecordPush("mails", mails.Count);
            await _notifications.Clients.All.SendAsync("MailsUpdated", mails);
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

        public async Task ReportCommandResult(OutlookCommandResult result)
        {
            var level = result.Success ? "info" : "warn";
            _addinStatus.AddLog(level, $"Command result {result.CommandId}: {result.Message}");
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
