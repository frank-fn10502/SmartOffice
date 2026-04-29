using Microsoft.AspNetCore.SignalR;
using Microsoft.Extensions.Options;
using SmartOffice.Hub.Hubs;
using SmartOffice.Hub.Models;

namespace SmartOffice.Hub.Services.MockAddins
{
    /// <summary>
    /// Development-only Outlook Add-in mock。
    /// 這個 worker 走與真實 Outlook Add-in 相同的 Hub 邊界：poll command、
    /// 將結果 push 進 Hub cache，接著透過 SignalR 通知 browser client。
    /// </summary>
    public class OutlookMockAddinWorker : BackgroundService
    {
        private readonly CommandQueue _commandQueue;
        private readonly MailStore _mailStore;
        private readonly ChatStore _chatStore;
        private readonly AddinStatusStore _addinStatus;
        private readonly IHubContext<NotificationHub> _hub;
        private readonly AddinMockOptions _options;
        private readonly ILogger<OutlookMockAddinWorker> _logger;
        private bool _seeded;

        public OutlookMockAddinWorker(
            CommandQueue commandQueue,
            MailStore mailStore,
            ChatStore chatStore,
            AddinStatusStore addinStatus,
            IHubContext<NotificationHub> hub,
            IOptions<AddinMockOptions> options,
            ILogger<OutlookMockAddinWorker> logger)
        {
            _commandQueue = commandQueue;
            _mailStore = mailStore;
            _chatStore = chatStore;
            _addinStatus = addinStatus;
            _hub = hub;
            _options = options.Value;
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken cancellationToken)
        {
            if (!_options.Enabled || !_options.Outlook.Enabled)
                return;

            _logger.LogInformation("Outlook mock Add-in is enabled.");
            await SeedInitialStateAsync(cancellationToken);

            while (!cancellationToken.IsCancellationRequested)
            {
                PendingCommand? command = null;

                try
                {
                    command = await _commandQueue.DequeueAsync(TimeSpan.FromSeconds(5), cancellationToken);
                    _addinStatus.RecordPoll(command?.Type);
                    await BroadcastStatusAsync(cancellationToken);

                    if (command == null)
                        continue;

                    await Task.Delay(_options.ResponseDelayMilliseconds, cancellationToken);
                    await HandleCommandAsync(command, cancellationToken);
                }
                catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                {
                    break;
                }
                catch (Exception ex)
                {
                    var commandType = command?.Type ?? "none";
                    _addinStatus.AddLog("error", $"Outlook mock failed while handling command '{commandType}': {ex.Message}");
                    await _hub.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs(), cancellationToken);
                    _logger.LogError(ex, "Outlook mock Add-in failed while handling command {CommandType}.", commandType);
                }
            }
        }

        private async Task SeedInitialStateAsync(CancellationToken cancellationToken)
        {
            if (_seeded)
                return;

            _seeded = true;

            var folders = CreateFolders();
            _mailStore.SetFolders(folders);
            _addinStatus.RecordPush("mock folders", folders.Count);
            _addinStatus.AddLog("info", "Outlook mock Add-in seeded folder cache.");

            _chatStore.Add(new ChatMessageDto
            {
                Source = "outlook",
                Text = "Outlook mock Add-in is connected. Web UI is reading Hub API data.",
                Timestamp = DateTime.Now
            });

            await _hub.Clients.All.SendAsync("FoldersUpdated", folders, cancellationToken);
            await _hub.Clients.All.SendAsync("AddinStatus", _addinStatus.GetStatus(), cancellationToken);
            await _hub.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs(), cancellationToken);
        }

        private async Task HandleCommandAsync(PendingCommand command, CancellationToken cancellationToken)
        {
            switch (command.Type)
            {
                case "fetch_folders":
                    await PushFoldersAsync(cancellationToken);
                    break;
                case "fetch_mails":
                    await PushMailsAsync(command.MailsRequest ?? new FetchMailsRequest(), cancellationToken);
                    break;
                default:
                    _addinStatus.AddLog("warn", $"Outlook mock received unsupported command: {command.Type}");
                    await _hub.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs(), cancellationToken);
                    break;
            }
        }

        private async Task PushFoldersAsync(CancellationToken cancellationToken)
        {
            var folders = CreateFolders();
            _mailStore.SetFolders(folders);
            _addinStatus.RecordPush("folders", folders.Count);

            await _hub.Clients.All.SendAsync("FoldersUpdated", folders, cancellationToken);
            await BroadcastStatusAsync(cancellationToken);
            await _hub.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs(), cancellationToken);
        }

        private async Task PushMailsAsync(FetchMailsRequest request, CancellationToken cancellationToken)
        {
            var mails = CreateMails(request);
            _mailStore.SetMails(mails);
            _addinStatus.RecordPush("mails", mails.Count);

            await _hub.Clients.All.SendAsync("MailsUpdated", mails, cancellationToken);
            await BroadcastStatusAsync(cancellationToken);
            await _hub.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs(), cancellationToken);
        }

        private async Task BroadcastStatusAsync(CancellationToken cancellationToken)
        {
            await _hub.Clients.All.SendAsync("AddinStatus", _addinStatus.GetStatus(), cancellationToken);
        }

        private static List<FolderDto> CreateFolders()
        {
            return new List<FolderDto>
            {
                new()
                {
                    Name = "Mailbox - Mock User",
                    FolderPath = "\\\\Mailbox - Mock User",
                    ItemCount = 42,
                    SubFolders =
                    {
                        new() { Name = "Inbox", FolderPath = "\\\\Mailbox - Mock User\\Inbox", ItemCount = 18 },
                        new() { Name = "Sent Items", FolderPath = "\\\\Mailbox - Mock User\\Sent Items", ItemCount = 9 },
                        new() { Name = "Drafts", FolderPath = "\\\\Mailbox - Mock User\\Drafts", ItemCount = 2 },
                        new()
                        {
                            Name = "Projects",
                            FolderPath = "\\\\Mailbox - Mock User\\Projects",
                            ItemCount = 13,
                            SubFolders =
                            {
                                new() { Name = "SmartOffice", FolderPath = "\\\\Mailbox - Mock User\\Projects\\SmartOffice", ItemCount = 7 },
                                new() { Name = "Vendors", FolderPath = "\\\\Mailbox - Mock User\\Projects\\Vendors", ItemCount = 6 }
                            }
                        },
                        new() { Name = "Archive", FolderPath = "\\\\Mailbox - Mock User\\Archive", ItemCount = 21 }
                    }
                }
            };
        }

        private static List<MailItemDto> CreateMails(FetchMailsRequest request)
        {
            var folderName = string.IsNullOrWhiteSpace(request.FolderPath)
                ? "Inbox"
                : request.FolderPath.Split('\\', StringSplitOptions.RemoveEmptyEntries).LastOrDefault() ?? "Inbox";
            var maxCount = Math.Clamp(request.MaxCount, 1, 100);
            var count = Math.Min(maxCount, 8);

            return Enumerable.Range(1, count)
                .Select(index =>
                {
                    var received = DateTime.Now.AddMinutes(-index * 37);
                    return new MailItemDto
                    {
                        Subject = $"[{folderName}] Mock mail #{index}: Hub protocol sample",
                        SenderName = index % 2 == 0 ? "SmartOffice Bot" : "Mock Project Lead",
                        SenderEmail = index % 2 == 0 ? "bot@local.smartoffice" : "lead@local.smartoffice",
                        ReceivedTime = received,
                        FolderPath = request.FolderPath,
                        Body = $"This is mock mail #{index} generated by SmartOffice.Hub.\n\nRange: {request.Range}\nFolder: {folderName}\nGenerated: {received:G}",
                        BodyHtml = $"""
                            <article style="font-family:Segoe UI,Arial,sans-serif;line-height:1.5;color:#1f2937">
                              <h2 style="margin:0 0 8px">Mock mail #{index}</h2>
                              <p>This message was generated by <strong>SmartOffice.Hub</strong>, not by frontend mock data.</p>
                              <ul>
                                <li>Range: {request.Range}</li>
                                <li>Folder: {folderName}</li>
                                <li>Generated: {received:G}</li>
                              </ul>
                            </article>
                            """
                    };
                })
                .ToList();
        }
    }
}
