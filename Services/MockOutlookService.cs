using Microsoft.AspNetCore.SignalR;
using SmartOffice.Hub.Hubs;
using SmartOffice.Hub.Models;

namespace SmartOffice.Hub.Services
{
    public class MockOutlookService
    {
        private readonly object _lock = new();
        private readonly IWebHostEnvironment _environment;
        private readonly MailStore _mailStore;
        private readonly ChatStore _chatStore;
        private readonly AddinStatusStore _addinStatus;
        private readonly IHubContext<NotificationHub> _notifications;
        private readonly List<MailItemDto> _mockMails = new();
        private List<FolderDto> _mockFolders = new();
        private List<OutlookCategoryDto> _mockCategories = new();
        private List<OutlookRuleDto> _mockRules = new();
        private List<CalendarEventDto> _mockCalendar = new();

        public MockOutlookService(
            IWebHostEnvironment environment,
            MailStore mailStore,
            ChatStore chatStore,
            AddinStatusStore addinStatus,
            IHubContext<NotificationHub> notifications)
        {
            _environment = environment;
            _mailStore = mailStore;
            _chatStore = chatStore;
            _addinStatus = addinStatus;
            _notifications = notifications;
        }

        public bool IsEnabled => _environment.IsEnvironment("Mock");

        public void Seed()
        {
            if (!IsEnabled) return;

            lock (_lock)
            {
                EnsureSeeded();
                _mailStore.SetFolders(CloneFolders(_mockFolders));
                _mailStore.SetMails(FilterMails(MockPaths.Inbox, "1w", 10));
                _mailStore.SetCategories(new List<OutlookCategoryDto>(_mockCategories));
                _mailStore.SetRules(new List<OutlookRuleDto>(_mockRules));
                _mailStore.SetCalendarEvents(new List<CalendarEventDto>(_mockCalendar));
                SeedChat();
                _addinStatus.RecordMockBackendReady();
            }
        }

        public async Task<bool> TryDispatchAsync(PendingCommand command, CancellationToken ct = default)
        {
            if (!IsEnabled) return false;

            List<FolderDto>? folders = null;
            List<MailItemDto>? mails = null;
            List<OutlookCategoryDto>? categories = null;
            List<OutlookRuleDto>? rules = null;
            List<CalendarEventDto>? calendar = null;
            var resultMessage = string.Empty;

            lock (_lock)
            {
                EnsureSeeded();
                switch (command.Type)
                {
                    case "fetch_folders":
                        folders = CloneFolders(_mockFolders);
                        _mailStore.SetFolders(folders);
                        break;
                    case "fetch_mails":
                        mails = FilterMails(
                            command.MailsRequest?.FolderPath ?? MockPaths.Inbox,
                            command.MailsRequest?.Range ?? "1d",
                            command.MailsRequest?.MaxCount ?? 10);
                        _mailStore.SetMails(mails);
                        break;
                    case "fetch_categories":
                        categories = new List<OutlookCategoryDto>(_mockCategories);
                        _mailStore.SetCategories(categories);
                        break;
                    case "fetch_rules":
                        rules = new List<OutlookRuleDto>(_mockRules);
                        _mailStore.SetRules(rules);
                        break;
                    case "fetch_calendar":
                        calendar = FilterCalendar(command.CalendarRequest?.DaysForward ?? 14);
                        _mailStore.SetCalendarEvents(calendar);
                        break;
                    case "mark_mail_read":
                        UpdateMailMarker(command.MailMarkerRequest, mail => mail.IsRead = true);
                        mails = SyncVisibleMails();
                        _mailStore.SetMails(mails);
                        break;
                    case "mark_mail_unread":
                        UpdateMailMarker(command.MailMarkerRequest, mail => mail.IsRead = false);
                        mails = SyncVisibleMails();
                        _mailStore.SetMails(mails);
                        break;
                    case "mark_mail_task":
                        UpdateMailMarker(command.MailMarkerRequest, mail =>
                        {
                            mail.IsMarkedAsTask = true;
                            mail.FlagInterval = "today";
                            mail.FlagRequest = "今天";
                            mail.TaskStartDate = DateTime.Now.Date;
                            mail.TaskDueDate = DateTime.Now.Date;
                            mail.TaskCompletedDate = null;
                            mail.Importance = "high";
                        });
                        mails = SyncVisibleMails();
                        _mailStore.SetMails(mails);
                        break;
                    case "clear_mail_task":
                        UpdateMailMarker(command.MailMarkerRequest, mail =>
                        {
                            mail.IsMarkedAsTask = false;
                            mail.FlagInterval = "none";
                            mail.FlagRequest = string.Empty;
                            mail.TaskStartDate = null;
                            mail.TaskDueDate = null;
                            mail.TaskCompletedDate = null;
                            mail.Importance = "normal";
                        });
                        mails = SyncVisibleMails();
                        _mailStore.SetMails(mails);
                        break;
                    case "set_mail_categories":
                        UpdateMailMarker(command.MailMarkerRequest, mail => mail.Categories = command.MailMarkerRequest?.Categories ?? string.Empty);
                        mails = SyncVisibleMails();
                        _mailStore.SetMails(mails);
                        break;
                    case "upsert_category":
                        UpsertCategory(command.CategoryRequest);
                        categories = new List<OutlookCategoryDto>(_mockCategories);
                        _mailStore.SetCategories(categories);
                        break;
                    case "update_mail_properties":
                        UpdateMailProperties(command.MailPropertiesRequest);
                        mails = SyncVisibleMails();
                        categories = new List<OutlookCategoryDto>(_mockCategories);
                        _mailStore.SetMails(mails);
                        _mailStore.SetCategories(categories);
                        break;
                    case "create_folder":
                        CreateFolder(command.CreateFolderRequest);
                        folders = CloneFolders(_mockFolders);
                        _mailStore.SetFolders(folders);
                        break;
                    case "delete_folder":
                        DeleteFolder(command.DeleteFolderRequest);
                        folders = CloneFolders(_mockFolders);
                        _mailStore.SetFolders(folders);
                        break;
                    case "move_mail":
                        MoveMail(command.MoveMailRequest);
                        mails = SyncVisibleMails(command.MoveMailRequest?.MailId);
                        folders = CloneFolders(_mockFolders);
                        _mailStore.SetMails(mails);
                        _mailStore.SetFolders(folders);
                        break;
                    case "ping":
                        break;
                    default:
                        resultMessage = $"Mock backend ignored unsupported command: {command.Type}";
                        break;
                }

                _addinStatus.RecordMockDispatch(command.Type);
                resultMessage = string.IsNullOrWhiteSpace(resultMessage)
                    ? $"{command.Type} completed by mock backend"
                    : resultMessage;
            }

            if (folders is not null) await _notifications.Clients.All.SendAsync("FoldersUpdated", folders, ct);
            if (mails is not null) await _notifications.Clients.All.SendAsync("MailsUpdated", mails, ct);
            if (categories is not null) await _notifications.Clients.All.SendAsync("CategoriesUpdated", categories, ct);
            if (rules is not null) await _notifications.Clients.All.SendAsync("RulesUpdated", rules, ct);
            if (calendar is not null) await _notifications.Clients.All.SendAsync("CalendarUpdated", calendar, ct);
            await _notifications.Clients.All.SendAsync("AddinStatus", _addinStatus.GetStatus(), ct);
            await _notifications.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs(), ct);
            await _notifications.Clients.All.SendAsync("CommandResult", new OutlookCommandResult
            {
                CommandId = command.Id,
                Success = true,
                Message = resultMessage,
                Timestamp = DateTime.Now,
            }, ct);
            return true;
        }

        public async Task<bool> TryReplyToChatAsync(ChatMessageDto message, CancellationToken ct = default)
        {
            if (!IsEnabled || !string.Equals(message.Source, "web", StringComparison.OrdinalIgnoreCase))
                return false;

            ChatMessageDto reply;
            lock (_lock)
            {
                EnsureSeeded();
                reply = new ChatMessageDto
                {
                    Source = "outlook",
                    Text = BuildChatReply(message.Text),
                    Timestamp = DateTime.Now,
                };
                _chatStore.Add(reply);
                _addinStatus.AddLog("info", "Mock Outlook chat reply generated");
            }

            await _notifications.Clients.All.SendAsync("NewChatMessage", reply, ct);
            await _notifications.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs(), ct);
            return true;
        }

        private void EnsureSeeded()
        {
            if (_mockFolders.Count > 0) return;

            _mockFolders = new List<FolderDto>
            {
                new()
                {
                    Name = "Mailbox - Mock User",
                    FolderPath = "\\\\Mailbox - Mock User",
                    ItemCount = 0,
                    SubFolders = new List<FolderDto>
                    {
                        Folder("Inbox", MockPaths.Inbox, 4, Folder("客戶專案", MockPaths.ClientProjects, 1)),
                        Folder("Sent Items", MockPaths.Sent, 1),
                        Folder("Archive", MockPaths.Archive, 2, Folder("2026 專案封存", MockPaths.Archive2026, 1)),
                        Folder("Drafts", MockPaths.Drafts, 1),
                        Folder("Deleted Items", MockPaths.Deleted, 0),
                    }
                }
            };

            var now = DateTime.Now;
            _mockMails.AddRange(new[]
            {
                Mail("mock-001", "週會議程與客戶需求整理", "Ada Chen", "ada.chen@example.test", now.AddMinutes(-28), MockPaths.Inbox, false, "客戶,待辦", true, "today", "今天"),
                Mail("mock-002", "Re: 合約附件確認", "Ben Lin", "ben.lin@example.test", now.AddHours(-2), MockPaths.Inbox, true, "", false, "none", ""),
                Mail("mock-003", "Office 2016 add-in hover 測試", "QA Lab", "qa@example.test", now.AddHours(-4), MockPaths.Inbox, false, "測試", false, "none", "", bodyHtml: ""),
                Mail("mock-004", "下週 demo 時程", "Chris Wang", "chris.wang@example.test", now.AddDays(-1), MockPaths.Inbox, true, "追蹤", true, "next_week", "下週"),
                Mail("mock-005", "專案資料夾歸檔樣本", "Dana Hsu", "dana.hsu@example.test", now.AddDays(-2), MockPaths.ClientProjects, true, "客戶", false, "none", ""),
                Mail("mock-006", "已寄出的測試郵件", "Mock User", "mock.user@example.test", now.AddDays(-3), MockPaths.Sent, true, "", false, "none", ""),
                Mail("mock-007", "一週前的封存通知", "System Notice", "notice@example.test", now.AddDays(-7), MockPaths.Archive, true, "測試", false, "none", ""),
                Mail("mock-008", "草稿：內部追蹤事項", "Mock User", "mock.user@example.test", now.AddDays(-10), MockPaths.Drafts, false, "待辦", true, "no_date", "Follow up"),
                Mail("mock-009", "上月客戶回覆", "Eve Huang", "eve.huang@example.test", now.AddDays(-25), MockPaths.Archive2026, true, "客戶", false, "none", ""),
            });

            _mockCategories = new List<OutlookCategoryDto>
            {
                new() { Name = "客戶", Color = "preset5", ShortcutKey = "" },
                new() { Name = "待辦", Color = "preset1", ShortcutKey = "" },
                new() { Name = "測試", Color = "preset4", ShortcutKey = "" },
                new() { Name = "追蹤", Color = "preset3", ShortcutKey = "" },
            };

            _mockRules = new List<OutlookRuleDto>
            {
                new()
                {
                    Name = "客戶郵件標記",
                    Enabled = true,
                    ExecutionOrder = 1,
                    Conditions = new List<string> { "sender contains example.test" },
                    Actions = new List<string> { "assign category 客戶" },
                },
                new()
                {
                    Name = "重要追蹤提醒",
                    Enabled = false,
                    ExecutionOrder = 2,
                    Conditions = new List<string> { "subject contains demo" },
                    Actions = new List<string> { "mark importance high", "flag for follow up" },
                    Exceptions = new List<string> { "sender is mock.user@example.test" },
                }
            };

            _mockCalendar = new List<CalendarEventDto>
            {
                new()
                {
                    Id = "mock-cal-001",
                    Subject = "SmartOffice mock sync review",
                    Start = now.Date.AddHours(15),
                    End = now.Date.AddHours(15).AddMinutes(30),
                    Location = "Teams",
                    Organizer = "mock.user@example.test",
                    RequiredAttendees = "ada.chen@example.test",
                    BusyStatus = "busy",
                },
                new()
                {
                    Id = "mock-cal-002",
                    Subject = "客戶需求釐清",
                    Start = now.Date.AddDays(2).AddHours(10),
                    End = now.Date.AddDays(2).AddHours(11),
                    Location = "會議室 3A",
                    Organizer = "ada.chen@example.test",
                    RequiredAttendees = "mock.user@example.test; dana.hsu@example.test",
                    IsRecurring = false,
                    BusyStatus = "tentative",
                },
                new()
                {
                    Id = "mock-cal-003",
                    Subject = "每週產品站會",
                    Start = now.Date.AddDays(6).AddHours(9),
                    End = now.Date.AddDays(6).AddHours(9).AddMinutes(45),
                    Location = "Teams",
                    Organizer = "mock.user@example.test",
                    RequiredAttendees = "product@example.test",
                    IsRecurring = true,
                    BusyStatus = "busy",
                }
            };
        }

        private void SeedChat()
        {
            if (_chatStore.GetAll().Count > 0) return;
            _chatStore.Add(new ChatMessageDto
            {
                Source = "outlook",
                Text = "Mock Outlook 已連線，可測試郵件、folder、category、calendar 與 chat 流程。",
                Timestamp = DateTime.Now.AddMinutes(-3),
            });
        }

        private List<MailItemDto> FilterMails(string folderPath, string range, int maxCount)
        {
            var target = string.IsNullOrWhiteSpace(folderPath) ? MockPaths.Inbox : folderPath;
            var since = RangeStart(range);
            return _mockMails
                .Where(mail => mail.FolderPath == target && mail.ReceivedTime >= since)
                .OrderByDescending(mail => mail.ReceivedTime)
                .Take(Math.Max(1, maxCount))
                .Select(CloneMail)
                .ToList();
        }

        private List<CalendarEventDto> FilterCalendar(int daysForward)
        {
            var start = DateTime.Now.Date;
            var end = start.AddDays(Math.Max(1, daysForward));
            return _mockCalendar
                .Where(item => item.Start >= start && item.Start < end)
                .OrderBy(item => item.Start)
                .Select(CloneCalendarEvent)
                .ToList();
        }

        private List<MailItemDto> SyncVisibleMails(string? removedMailId = null)
        {
            var current = _mailStore.GetMails();
            var ids = current
                .Where(mail => mail.Id != removedMailId)
                .Select(mail => mail.Id)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            return _mockMails
                .Where(mail => ids.Contains(mail.Id))
                .OrderByDescending(mail => mail.ReceivedTime)
                .Select(CloneMail)
                .ToList();
        }

        private void UpsertCategory(CategoryCommandRequest? request)
        {
            if (request is null || string.IsNullOrWhiteSpace(request.Name)) return;
            var existing = _mockCategories.FirstOrDefault(category => category.Name.Equals(request.Name, StringComparison.OrdinalIgnoreCase));
            if (existing is null)
            {
                _mockCategories.Add(new OutlookCategoryDto
                {
                    Name = request.Name.Trim(),
                    Color = string.IsNullOrWhiteSpace(request.Color) ? "preset0" : request.Color,
                    ShortcutKey = request.ShortcutKey,
                });
                return;
            }

            existing.Color = string.IsNullOrWhiteSpace(request.Color) ? existing.Color : request.Color;
            existing.ShortcutKey = request.ShortcutKey;
        }

        private void UpdateMailMarker(MailMarkerCommandRequest? request, Action<MailItemDto> update)
        {
            if (request is null || string.IsNullOrWhiteSpace(request.MailId)) return;
            var mail = _mockMails.FirstOrDefault(item => item.Id == request.MailId);
            if (mail is null) return;
            update(mail);
        }

        private void UpdateMailProperties(MailPropertiesCommandRequest? request)
        {
            if (request is null) return;
            var mail = _mockMails.FirstOrDefault(item => item.Id == request.MailId);
            if (mail is null) return;

            if (request.IsRead.HasValue) mail.IsRead = request.IsRead.Value;
            ApplyFlag(mail, request);
            mail.Categories = string.Join(",", request.Categories.Where(category => !string.IsNullOrWhiteSpace(category)).Select(category => category.Trim()));

            foreach (var category in request.NewCategories)
                UpsertCategory(new CategoryCommandRequest { Name = category.Name, Color = category.Color, ShortcutKey = category.ShortcutKey });
        }

        private void CreateFolder(CreateFolderRequest? request)
        {
            if (request is null || string.IsNullOrWhiteSpace(request.ParentFolderPath) || string.IsNullOrWhiteSpace(request.Name)) return;
            var parent = FindFolder(_mockFolders, request.ParentFolderPath);
            if (parent is null) return;
            var name = request.Name.Trim();
            if (parent.SubFolders.Any(folder => folder.Name.Equals(name, StringComparison.OrdinalIgnoreCase))) return;
            var path = $"{request.ParentFolderPath}\\{name}";
            parent.SubFolders.Add(Folder(name, path, 0));
        }

        private void DeleteFolder(DeleteFolderRequest? request)
        {
            if (request is null || string.IsNullOrWhiteSpace(request.FolderPath)) return;
            DeleteFolderFrom(_mockFolders, request.FolderPath);
            _mockMails.RemoveAll(mail => mail.FolderPath.StartsWith(request.FolderPath, StringComparison.OrdinalIgnoreCase));
            RefreshFolderCounts();
        }

        private void MoveMail(MoveMailRequest? request)
        {
            if (request is null || string.IsNullOrWhiteSpace(request.MailId) || string.IsNullOrWhiteSpace(request.DestinationFolderPath)) return;
            var mail = _mockMails.FirstOrDefault(item => item.Id == request.MailId);
            if (mail is null || mail.FolderPath == request.DestinationFolderPath) return;
            mail.FolderPath = request.DestinationFolderPath;
            RefreshFolderCounts();
        }

        private static void ApplyFlag(MailItemDto mail, MailPropertiesCommandRequest request)
        {
            mail.FlagInterval = string.IsNullOrWhiteSpace(request.FlagInterval) ? "none" : request.FlagInterval;
            mail.FlagRequest = request.FlagRequest;
            mail.IsMarkedAsTask = mail.FlagInterval != "none";
            mail.TaskCompletedDate = mail.FlagInterval == "complete" ? request.TaskCompletedDate ?? DateTime.Now.Date : null;

            if (mail.FlagInterval == "custom")
            {
                mail.TaskStartDate = request.TaskStartDate;
                mail.TaskDueDate = request.TaskDueDate;
            }
            else if (mail.IsMarkedAsTask && mail.FlagInterval != "complete")
            {
                var due = FlagDueDate(mail.FlagInterval);
                mail.TaskStartDate = DateTime.Now.Date;
                mail.TaskDueDate = due;
            }
            else
            {
                mail.TaskStartDate = null;
                mail.TaskDueDate = null;
            }

            mail.Importance = mail.IsMarkedAsTask ? "high" : "normal";
        }

        private static DateTime? FlagDueDate(string flagInterval)
        {
            var today = DateTime.Now.Date;
            return flagInterval switch
            {
                "today" => today,
                "tomorrow" => today.AddDays(1),
                "this_week" => today.AddDays(5),
                "next_week" => today.AddDays(7),
                "no_date" => null,
                _ => null,
            };
        }

        private static DateTime RangeStart(string range)
        {
            var now = DateTime.Now;
            return range switch
            {
                "1w" => now.AddDays(-7),
                "1m" => now.AddMonths(-1),
                _ => now.AddDays(-1),
            };
        }

        private static string BuildChatReply(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "Mock Outlook 收到空白訊息。";

            return $"Mock Outlook 已收到：「{text.Trim()}」。目前本機資料可直接測試 request、更新與即時廣播。";
        }

        private void RefreshFolderCounts()
        {
            foreach (var folder in FlattenFolders(_mockFolders))
                folder.ItemCount = _mockMails.Count(mail => mail.FolderPath == folder.FolderPath);
        }

        private static FolderDto Folder(string name, string folderPath, int itemCount, params FolderDto[] children)
        {
            return new FolderDto
            {
                Name = name,
                FolderPath = folderPath,
                ItemCount = itemCount,
                SubFolders = children.ToList(),
            };
        }

        private static MailItemDto Mail(
            string id,
            string subject,
            string senderName,
            string senderEmail,
            DateTime receivedTime,
            string folderPath,
            bool isRead,
            string categories,
            bool isMarkedAsTask,
            string flagInterval,
            string flagRequest,
            string? bodyHtml = null)
        {
            var body = $"Mock 郵件內容：{subject}\n\n這封郵件用於本機測試 Web UI、drag/drop 與 contract 行為。";
            return new MailItemDto
            {
                Id = id,
                Subject = subject,
                SenderName = senderName,
                SenderEmail = senderEmail,
                ReceivedTime = receivedTime,
                Body = body,
                BodyHtml = bodyHtml ?? $"<article><h2>{subject}</h2><p>Mock 郵件內容，用於本機測試 Web UI 與 Outlook contract。</p></article>",
                FolderPath = folderPath,
                Categories = categories,
                IsRead = isRead,
                IsMarkedAsTask = isMarkedAsTask,
                FlagInterval = flagInterval,
                FlagRequest = flagRequest,
                TaskDueDate = isMarkedAsTask ? DateTime.Now.Date.AddDays(1) : null,
                Importance = isMarkedAsTask ? "high" : "normal",
                Sensitivity = "normal",
            };
        }

        private static FolderDto? FindFolder(List<FolderDto> folders, string path)
        {
            foreach (var folder in folders)
            {
                if (folder.FolderPath == path) return folder;
                var child = FindFolder(folder.SubFolders, path);
                if (child is not null) return child;
            }

            return null;
        }

        private static bool DeleteFolderFrom(List<FolderDto> folders, string path)
        {
            var removed = folders.RemoveAll(folder => folder.FolderPath == path) > 0;
            if (removed) return true;
            return folders.Any(folder => DeleteFolderFrom(folder.SubFolders, path));
        }

        private static IEnumerable<FolderDto> FlattenFolders(List<FolderDto> folders)
        {
            foreach (var folder in folders)
            {
                yield return folder;
                foreach (var child in FlattenFolders(folder.SubFolders))
                    yield return child;
            }
        }

        private static List<FolderDto> CloneFolders(List<FolderDto> folders)
        {
            return folders.Select(folder => new FolderDto
            {
                Name = folder.Name,
                FolderPath = folder.FolderPath,
                ItemCount = folder.ItemCount,
                SubFolders = CloneFolders(folder.SubFolders),
            }).ToList();
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

        private static CalendarEventDto CloneCalendarEvent(CalendarEventDto item)
        {
            return new CalendarEventDto
            {
                Id = item.Id,
                Subject = item.Subject,
                Start = item.Start,
                End = item.End,
                Location = item.Location,
                Organizer = item.Organizer,
                RequiredAttendees = item.RequiredAttendees,
                IsRecurring = item.IsRecurring,
                BusyStatus = item.BusyStatus,
            };
        }

        private static class MockPaths
        {
            public const string Inbox = "\\\\Mailbox - Mock User\\Inbox";
            public const string ClientProjects = "\\\\Mailbox - Mock User\\Inbox\\客戶專案";
            public const string Sent = "\\\\Mailbox - Mock User\\Sent Items";
            public const string Archive = "\\\\Mailbox - Mock User\\Archive";
            public const string Archive2026 = "\\\\Mailbox - Mock User\\Archive\\2026 專案封存";
            public const string Drafts = "\\\\Mailbox - Mock User\\Drafts";
            public const string Deleted = "\\\\Mailbox - Mock User\\Deleted Items";
        }
    }
}
