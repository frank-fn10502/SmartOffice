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
            var rules = CreateRules();
            var categories = CreateCategories();
            var events = CreateCalendarEvents(new FetchCalendarRequest());
            _mailStore.SetFolders(folders);
            _mailStore.SetRules(rules);
            _mailStore.SetCategories(categories);
            _mailStore.SetCalendarEvents(events);
            _addinStatus.RecordPush("mock folders", folders.Count);
            _addinStatus.RecordPush("mock rules", rules.Count);
            _addinStatus.RecordPush("mock categories", categories.Count);
            _addinStatus.RecordPush("mock calendar", events.Count);
            _addinStatus.AddLog("info", "Outlook mock Add-in seeded folder cache.");

            _chatStore.Add(new ChatMessageDto
            {
                Source = "outlook",
                Text = "Outlook mock Add-in is connected. Web UI is reading Hub API data.",
                Timestamp = DateTime.Now
            });

            await _hub.Clients.All.SendAsync("FoldersUpdated", folders, cancellationToken);
            await _hub.Clients.All.SendAsync("RulesUpdated", rules, cancellationToken);
            await _hub.Clients.All.SendAsync("CategoriesUpdated", categories, cancellationToken);
            await _hub.Clients.All.SendAsync("CalendarUpdated", events, cancellationToken);
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
                case "fetch_rules":
                    await PushRulesAsync(cancellationToken);
                    break;
                case "fetch_categories":
                    await PushCategoriesAsync(cancellationToken);
                    break;
                case "fetch_calendar":
                    await PushCalendarAsync(command.CalendarRequest ?? new FetchCalendarRequest(), cancellationToken);
                    break;
                case "update_mail_properties":
                    await ApplyMailPropertiesAsync(command.MailPropertiesRequest, cancellationToken);
                    break;
                case "upsert_category":
                    await ApplyUpsertCategoryAsync(command.CategoryRequest, cancellationToken);
                    break;
                case "mark_mail_read":
                case "mark_mail_unread":
                case "mark_mail_task":
                case "clear_mail_task":
                case "set_mail_categories":
                    await ApplyMailMarkerAsync(command.Type, command.MailMarkerRequest, cancellationToken);
                    break;
                case "create_folder":
                    await ApplyCreateFolderAsync(command.CreateFolderRequest, cancellationToken);
                    break;
                case "delete_folder":
                    await ApplyDeleteFolderAsync(command.DeleteFolderRequest, cancellationToken);
                    break;
                case "move_mail":
                    await ApplyMoveMailAsync(command.MoveMailRequest, cancellationToken);
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

        private async Task PushRulesAsync(CancellationToken cancellationToken)
        {
            var rules = CreateRules();
            _mailStore.SetRules(rules);
            _addinStatus.RecordPush("rules", rules.Count);

            await _hub.Clients.All.SendAsync("RulesUpdated", rules, cancellationToken);
            await BroadcastStatusAsync(cancellationToken);
            await _hub.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs(), cancellationToken);
        }

        private async Task PushCategoriesAsync(CancellationToken cancellationToken)
        {
            var categories = _mailStore.GetCategories();
            if (categories.Count == 0)
                categories = CreateCategories();
            _mailStore.SetCategories(categories);
            _addinStatus.RecordPush("categories", categories.Count);

            await _hub.Clients.All.SendAsync("CategoriesUpdated", categories, cancellationToken);
            await BroadcastStatusAsync(cancellationToken);
            await _hub.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs(), cancellationToken);
        }

        private async Task PushCalendarAsync(FetchCalendarRequest request, CancellationToken cancellationToken)
        {
            var events = CreateCalendarEvents(request);
            _mailStore.SetCalendarEvents(events);
            _addinStatus.RecordPush("calendar", events.Count);

            await _hub.Clients.All.SendAsync("CalendarUpdated", events, cancellationToken);
            await BroadcastStatusAsync(cancellationToken);
            await _hub.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs(), cancellationToken);
        }

        private async Task ApplyMailMarkerAsync(string commandType, MailMarkerCommandRequest? request, CancellationToken cancellationToken)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.MailId))
            {
                await LogMockWarningAsync($"{commandType} ignored: missing mail id", cancellationToken);
                return;
            }

            var mails = _mailStore.GetMails();
            var mail = mails.FirstOrDefault(item => item.Id == request.MailId);
            if (mail == null)
            {
                await LogMockWarningAsync($"{commandType} ignored: mock mail not found", cancellationToken);
                return;
            }

            switch (commandType)
            {
                case "mark_mail_read":
                    mail.IsRead = true;
                    break;
                case "mark_mail_unread":
                    mail.IsRead = false;
                    break;
                case "mark_mail_task":
                    mail.IsMarkedAsTask = true;
                    mail.FlagInterval = "today";
                    mail.FlagRequest = FlagIntervalLabel(mail.FlagInterval);
                    mail.TaskStartDate = DateTime.Today;
                    mail.TaskDueDate = DateTime.Today;
                    mail.TaskCompletedDate = null;
                    break;
                case "clear_mail_task":
                    mail.IsMarkedAsTask = false;
                    mail.FlagInterval = "none";
                    mail.FlagRequest = string.Empty;
                    mail.TaskStartDate = null;
                    mail.TaskDueDate = null;
                    mail.TaskCompletedDate = null;
                    break;
                case "set_mail_categories":
                    mail.Categories = request.Categories;
                    break;
            }

            _mailStore.SetMails(mails);
            _addinStatus.RecordPush(commandType, 1);
            await _hub.Clients.All.SendAsync("MailsUpdated", mails, cancellationToken);
            await BroadcastStatusAsync(cancellationToken);
            await _hub.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs(), cancellationToken);
        }

        private async Task ApplyMailPropertiesAsync(MailPropertiesCommandRequest? request, CancellationToken cancellationToken)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.MailId))
            {
                await LogMockWarningAsync("update_mail_properties ignored: missing mail id", cancellationToken);
                return;
            }

            var mails = _mailStore.GetMails();
            var mail = mails.FirstOrDefault(item => item.Id == request.MailId);
            if (mail == null)
            {
                await LogMockWarningAsync("update_mail_properties ignored: mock mail not found", cancellationToken);
                return;
            }

            if (request.IsRead.HasValue)
                mail.IsRead = request.IsRead.Value;

            ApplyTaskFlag(mail, request);
            mail.Categories = string.Join(", ", request.Categories.Where(category => !string.IsNullOrWhiteSpace(category)).Distinct(StringComparer.OrdinalIgnoreCase));

            var categories = _mailStore.GetCategories();
            foreach (var category in request.NewCategories.Where(item => !string.IsNullOrWhiteSpace(item.Name)))
            {
                if (categories.Any(item => string.Equals(item.Name, category.Name, StringComparison.OrdinalIgnoreCase)))
                    continue;

                categories.Add(new OutlookCategoryDto
                {
                    Name = category.Name.Trim(),
                    Color = string.IsNullOrWhiteSpace(category.Color) ? "preset0" : category.Color,
                    ShortcutKey = category.ShortcutKey
                });
            }

            _mailStore.SetCategories(categories);
            _mailStore.SetMails(mails);
            _addinStatus.RecordPush("update_mail_properties", 1);
            await _hub.Clients.All.SendAsync("CategoriesUpdated", categories, cancellationToken);
            await _hub.Clients.All.SendAsync("MailsUpdated", mails, cancellationToken);
            await BroadcastStatusAsync(cancellationToken);
            await _hub.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs(), cancellationToken);
        }

        private async Task ApplyUpsertCategoryAsync(CategoryCommandRequest? request, CancellationToken cancellationToken)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Name))
            {
                await LogMockWarningAsync("upsert_category ignored: missing category name", cancellationToken);
                return;
            }

            var categoryName = request.Name.Trim();
            var categories = _mailStore.GetCategories();
            var existing = categories.FirstOrDefault(item => string.Equals(item.Name, categoryName, StringComparison.OrdinalIgnoreCase));

            if (existing == null)
            {
                categories.Add(new OutlookCategoryDto
                {
                    Name = categoryName,
                    Color = string.IsNullOrWhiteSpace(request.Color) ? "preset0" : request.Color,
                    ShortcutKey = request.ShortcutKey
                });
            }
            else
            {
                existing.Color = string.IsNullOrWhiteSpace(request.Color) ? existing.Color : request.Color;
                existing.ShortcutKey = request.ShortcutKey;
            }

            _mailStore.SetCategories(categories);
            _addinStatus.RecordPush("upsert_category", 1);
            await _hub.Clients.All.SendAsync("CategoriesUpdated", categories, cancellationToken);
            await BroadcastStatusAsync(cancellationToken);
            await _hub.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs(), cancellationToken);
        }

        private static void ApplyTaskFlag(MailItemDto mail, MailPropertiesCommandRequest request)
        {
            mail.FlagInterval = request.FlagInterval;
            mail.FlagRequest = request.FlagRequest.Trim();
            mail.TaskStartDate = request.TaskStartDate;
            mail.TaskDueDate = request.TaskDueDate;
            mail.TaskCompletedDate = request.TaskCompletedDate;
            mail.IsMarkedAsTask = request.FlagInterval != "none";

            if (request.FlagInterval == "none")
            {
                mail.FlagRequest = string.Empty;
                mail.TaskStartDate = null;
                mail.TaskDueDate = null;
                mail.TaskCompletedDate = null;
                return;
            }

            if (string.IsNullOrWhiteSpace(mail.FlagRequest))
                mail.FlagRequest = FlagIntervalLabel(request.FlagInterval);

            var today = DateTime.Today;
            switch (request.FlagInterval)
            {
                case "today":
                    mail.TaskStartDate = today;
                    mail.TaskDueDate = today;
                    mail.TaskCompletedDate = null;
                    break;
                case "tomorrow":
                    mail.TaskStartDate = today.AddDays(1);
                    mail.TaskDueDate = today.AddDays(1);
                    mail.TaskCompletedDate = null;
                    break;
                case "this_week":
                    mail.TaskStartDate = today;
                    mail.TaskDueDate = today.AddDays(2);
                    mail.TaskCompletedDate = null;
                    break;
                case "next_week":
                    mail.TaskStartDate = today.AddDays(7);
                    mail.TaskDueDate = today.AddDays(11);
                    mail.TaskCompletedDate = null;
                    break;
                case "no_date":
                    mail.TaskStartDate = null;
                    mail.TaskDueDate = null;
                    mail.TaskCompletedDate = null;
                    break;
                case "custom":
                    mail.TaskStartDate = request.TaskStartDate;
                    mail.TaskDueDate = request.TaskDueDate;
                    mail.TaskCompletedDate = null;
                    break;
                case "complete":
                    mail.TaskCompletedDate = request.TaskCompletedDate ?? DateTime.Now;
                    break;
            }
        }

        private static string FlagIntervalLabel(string interval)
        {
            return interval switch
            {
                "today" => "今天",
                "tomorrow" => "明天",
                "this_week" => "本週",
                "next_week" => "下週",
                "no_date" => "無日期",
                "custom" => "自訂日期",
                "complete" => "標示完成",
                _ => string.Empty
            };
        }

        private async Task ApplyCreateFolderAsync(CreateFolderRequest? request, CancellationToken cancellationToken)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.ParentFolderPath) || string.IsNullOrWhiteSpace(request.Name))
            {
                await LogMockWarningAsync("create_folder ignored: missing parent path or name", cancellationToken);
                return;
            }

            var folders = _mailStore.GetFolders();
            var parent = FindFolder(folders, request.ParentFolderPath);
            if (parent == null)
            {
                await LogMockWarningAsync("create_folder ignored: parent folder not found", cancellationToken);
                return;
            }

            if (parent.SubFolders.Any(folder => string.Equals(folder.Name, request.Name, StringComparison.OrdinalIgnoreCase)))
            {
                await LogMockWarningAsync("create_folder ignored: folder already exists", cancellationToken);
                return;
            }

            parent.SubFolders.Add(new FolderDto
            {
                Name = request.Name,
                FolderPath = $"{request.ParentFolderPath}\\{request.Name}",
                ItemCount = 0
            });
            _mailStore.SetFolders(folders);
            _addinStatus.RecordPush("create_folder", 1);
            await _hub.Clients.All.SendAsync("FoldersUpdated", folders, cancellationToken);
            await BroadcastStatusAsync(cancellationToken);
            await _hub.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs(), cancellationToken);
        }

        private async Task ApplyDeleteFolderAsync(DeleteFolderRequest? request, CancellationToken cancellationToken)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.FolderPath))
            {
                await LogMockWarningAsync("delete_folder ignored: missing folder path", cancellationToken);
                return;
            }

            var folders = _mailStore.GetFolders();
            if (!RemoveFolder(folders, request.FolderPath))
            {
                await LogMockWarningAsync("delete_folder ignored: folder not found or root folder selected", cancellationToken);
                return;
            }

            var childPathPrefix = $"{request.FolderPath}\\";
            var mails = _mailStore.GetMails()
                .Where(mail =>
                    !string.Equals(mail.FolderPath, request.FolderPath, StringComparison.OrdinalIgnoreCase)
                    && !mail.FolderPath.StartsWith(childPathPrefix, StringComparison.OrdinalIgnoreCase))
                .ToList();
            _mailStore.SetFolders(folders);
            _mailStore.SetMails(mails);
            _addinStatus.RecordPush("delete_folder", 1);
            await _hub.Clients.All.SendAsync("FoldersUpdated", folders, cancellationToken);
            await _hub.Clients.All.SendAsync("MailsUpdated", mails, cancellationToken);
            await BroadcastStatusAsync(cancellationToken);
            await _hub.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs(), cancellationToken);
        }

        private async Task ApplyMoveMailAsync(MoveMailRequest? request, CancellationToken cancellationToken)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.MailId) || string.IsNullOrWhiteSpace(request.DestinationFolderPath))
            {
                await LogMockWarningAsync("move_mail ignored: missing mail id or destination", cancellationToken);
                return;
            }

            var folders = _mailStore.GetFolders();
            if (FindFolder(folders, request.DestinationFolderPath) == null)
            {
                await LogMockWarningAsync("move_mail ignored: destination folder not found", cancellationToken);
                return;
            }

            var mails = _mailStore.GetMails();
            var mail = mails.FirstOrDefault(item => item.Id == request.MailId);
            if (mail == null)
            {
                await LogMockWarningAsync("move_mail ignored: mock mail not found", cancellationToken);
                return;
            }

            mail.FolderPath = request.DestinationFolderPath;
            _mailStore.SetMails(mails);
            _addinStatus.RecordPush("move_mail", 1);
            await _hub.Clients.All.SendAsync("MailsUpdated", mails, cancellationToken);
            await BroadcastStatusAsync(cancellationToken);
            await _hub.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs(), cancellationToken);
        }

        private async Task LogMockWarningAsync(string message, CancellationToken cancellationToken)
        {
            _addinStatus.AddLog("warn", message);
            await _hub.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs(), cancellationToken);
        }

        private static FolderDto? FindFolder(IEnumerable<FolderDto> folders, string folderPath)
        {
            foreach (var folder in folders)
            {
                if (string.Equals(folder.FolderPath, folderPath, StringComparison.OrdinalIgnoreCase))
                    return folder;

                var child = FindFolder(folder.SubFolders, folderPath);
                if (child != null)
                    return child;
            }

            return null;
        }

        private static bool RemoveFolder(List<FolderDto> folders, string folderPath)
        {
            for (var index = 0; index < folders.Count; index++)
            {
                var folder = folders[index];
                if (string.Equals(folder.FolderPath, folderPath, StringComparison.OrdinalIgnoreCase))
                {
                    folders.RemoveAt(index);
                    return true;
                }

                if (RemoveFolder(folder.SubFolders, folderPath))
                    return true;
            }

            return false;
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
                    var isFlagged = index % 3 == 0;
                    return new MailItemDto
                    {
                        Subject = $"[{folderName}] Mock mail #{index}: Hub protocol sample",
                        Id = $"mock-mail-{folderName}-{index}",
                        SenderName = index % 2 == 0 ? "SmartOffice Bot" : "Mock Project Lead",
                        SenderEmail = index % 2 == 0 ? "bot@local.smartoffice" : "lead@local.smartoffice",
                        ReceivedTime = received,
                        FolderPath = request.FolderPath,
                        Categories = index % 3 == 0 ? "Project, Follow-up" : index % 2 == 0 ? "Automation" : "",
                        IsRead = index % 2 == 0,
                        IsMarkedAsTask = isFlagged,
                        FlagInterval = isFlagged ? (index % 2 == 0 ? "tomorrow" : "today") : "none",
                        FlagRequest = isFlagged ? FlagIntervalLabel(index % 2 == 0 ? "tomorrow" : "today") : string.Empty,
                        TaskStartDate = isFlagged ? DateTime.Today : null,
                        TaskDueDate = isFlagged ? DateTime.Today.AddDays(index % 2) : null,
                        TaskCompletedDate = null,
                        Importance = index % 5 == 0 ? "high" : "normal",
                        Sensitivity = "normal",
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

        private static List<OutlookRuleDto> CreateRules()
        {
            return new List<OutlookRuleDto>
            {
                new()
                {
                    Name = "Project mail to SmartOffice",
                    Enabled = true,
                    ExecutionOrder = 1,
                    RuleType = "receive",
                    Conditions = { "subject contains SmartOffice", "sender domain is local.smartoffice" },
                    Actions = { "move to \\\\Mailbox - Mock User\\Projects\\SmartOffice", "assign category Project" }
                },
                new()
                {
                    Name = "Vendor invoices",
                    Enabled = true,
                    ExecutionOrder = 2,
                    RuleType = "receive",
                    Conditions = { "subject contains invoice", "has attachment" },
                    Actions = { "assign category Finance", "flag for follow-up" },
                    Exceptions = { "sender is internal finance" }
                },
                new()
                {
                    Name = "Low priority newsletters",
                    Enabled = false,
                    ExecutionOrder = 3,
                    RuleType = "receive",
                    Conditions = { "sender contains newsletter" },
                    Actions = { "move to \\\\Mailbox - Mock User\\Archive" }
                }
            };
        }

        private static List<OutlookCategoryDto> CreateCategories()
        {
            return new List<OutlookCategoryDto>
            {
                new() { Name = "Project", Color = "preset4" },
                new() { Name = "Follow-up", Color = "preset6" },
                new() { Name = "Automation", Color = "preset5" },
                new() { Name = "Finance", Color = "preset3" },
                new() { Name = "Reviewed", Color = "preset8" }
            };
        }

        private static List<CalendarEventDto> CreateCalendarEvents(FetchCalendarRequest request)
        {
            var daysForward = Math.Clamp(request.DaysForward, 1, 60);
            var now = DateTime.Now;
            var events = new List<CalendarEventDto>
            {
                new()
                {
                    Id = "mock-calendar-1",
                    Subject = "SmartOffice planning sync",
                    Start = now.Date.AddDays(1).AddHours(10),
                    End = now.Date.AddDays(1).AddHours(11),
                    Location = "Teams",
                    Organizer = "Mock Project Lead",
                    RequiredAttendees = "SmartOffice Team",
                    IsRecurring = true,
                    BusyStatus = "busy"
                },
                new()
                {
                    Id = "mock-calendar-2",
                    Subject = "Vendor contract review",
                    Start = now.Date.AddDays(3).AddHours(14),
                    End = now.Date.AddDays(3).AddHours(15).AddMinutes(30),
                    Location = "Meeting Room B",
                    Organizer = "Procurement",
                    RequiredAttendees = "Legal; Project Lead",
                    IsRecurring = false,
                    BusyStatus = "tentative"
                },
                new()
                {
                    Id = "mock-calendar-3",
                    Subject = "Customer follow-up",
                    Start = now.Date.AddDays(8).AddHours(9),
                    End = now.Date.AddDays(8).AddHours(9).AddMinutes(45),
                    Location = "Phone",
                    Organizer = "Account Manager",
                    RequiredAttendees = "Customer A; Account Manager",
                    IsRecurring = false,
                    BusyStatus = "busy"
                }
            };

            return events.Where(item => item.Start <= now.Date.AddDays(daysForward + 1)).ToList();
        }
    }
}
