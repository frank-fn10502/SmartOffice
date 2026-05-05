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
        private readonly AttachmentExportService _attachmentExports;
        private readonly IHubContext<NotificationHub> _notifications;
        private readonly List<MailItemDto> _mockMails = new();
        private List<FolderDto> _mockFolders = new();
        private List<OutlookStoreDto> _mockStores = new();
        private List<OutlookCategoryDto> _mockCategories = new();
        private List<OutlookRuleDto> _mockRules = new();
        private List<CalendarEventDto> _mockCalendar = new();

        public MockOutlookService(
            IWebHostEnvironment environment,
            MailStore mailStore,
            ChatStore chatStore,
            AddinStatusStore addinStatus,
            AttachmentExportService attachmentExports,
            IHubContext<NotificationHub> notifications)
        {
            _environment = environment;
            _mailStore = mailStore;
            _chatStore = chatStore;
            _addinStatus = addinStatus;
            _attachmentExports = attachmentExports;
            _notifications = notifications;
        }

        public bool IsEnabled => _environment.IsEnvironment("Mock");

        public void Seed()
        {
            if (!IsEnabled) return;

            lock (_lock)
            {
                EnsureSeeded();
                _mailStore.ApplyFolderBatch(BuildFolderRootsBatch(reset: true));
                _mailStore.SetMails(MockOutlookMailSearch.FilterMails(_mockMails, MockOutlookPaths.Inbox, MockOutlookPaths.Inbox, "1m", 30));
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

            FolderSyncBatchDto? folderBatch = null;
            MailSearchSliceResultDto? mailSearchSliceResult = null;
            MailSearchCompleteDto? mailSearchComplete = null;
            List<MailItemDto>? mails = null;
            MailItemDto? mail = null;
            MailBodyDto? mailBody = null;
            MailAttachmentsDto? mailAttachments = null;
            ExportedMailAttachmentDto? exportedAttachment = null;
            List<OutlookCategoryDto>? categories = null;
            List<OutlookRuleDto>? rules = null;
            List<CalendarEventDto>? calendar = null;
            var resultMessage = string.Empty;

            lock (_lock)
            {
                EnsureSeeded();
                switch (command.Type)
                {
                    case "fetch_folder_roots":
                        folderBatch = BuildFolderRootsBatch(command.FolderDiscoveryRequest?.Reset ?? true);
                        _mailStore.ApplyFolderBatch(folderBatch);
                        break;
                    case "fetch_folder_children":
                        folderBatch = BuildFolderChildrenBatch(command.FolderDiscoveryRequest);
                        _mailStore.ApplyFolderBatch(folderBatch);
                        break;
                    case "fetch_mails":
                        mails = MockOutlookMailSearch.FilterMails(
                            _mockMails,
                            MockOutlookPaths.Inbox,
                            command.MailsRequest?.FolderPath ?? MockOutlookPaths.Inbox,
                            command.MailsRequest?.Range ?? "1m",
                            command.MailsRequest?.MaxCount ?? 30);
                        _mailStore.SetMails(mails);
                        break;
                    case "fetch_mail_search_slice":
                        var request = command.MailSearchSliceRequest ?? new MailSearchSliceRequest();
                        var searchResults = MockOutlookMailSearch.FetchMailSearchSlice(_mockMails, request);
                        mailSearchSliceResult = new MailSearchSliceResultDto
                        {
                            SearchId = request.SearchId,
                            CommandId = command.Id,
                            ParentCommandId = request.ParentCommandId,
                            Sequence = request.SliceIndex + 1,
                            SliceIndex = request.SliceIndex,
                            SliceCount = request.SliceCount,
                            Reset = request.ResetSearchResults,
                            IsFinal = request.CompleteSearchOnSlice,
                            IsSliceComplete = true,
                            Mails = searchResults,
                            Message = "Mock mail search completed",
                        };
                        _mailStore.BeginMailSearch(mailSearchSliceResult.Reset);
                        _mailStore.ApplyMailSearchSliceResult(mailSearchSliceResult);
                        if (request.CompleteSearchOnSlice)
                        {
                            mailSearchComplete = new MailSearchCompleteDto
                            {
                                SearchId = mailSearchSliceResult.SearchId,
                                CommandId = command.Id,
                                ParentCommandId = request.ParentCommandId,
                                TotalCount = _mailStore.GetMailSearchResults().Count,
                                Message = "Mock mail search completed",
                            Timestamp = DateTime.Now,
                        };
                        }
                        break;
                    case "fetch_mail_body":
                        mailBody = FetchMailBody(command.MailBodyRequest);
                        if (mailBody is not null) _mailStore.UpdateMailBody(mailBody);
                        break;
                    case "fetch_mail_attachments":
                        mailAttachments = FetchMailAttachments(command.MailAttachmentsRequest);
                        if (mailAttachments is not null) _mailStore.SetMailAttachments(mailAttachments);
                        break;
                    case "export_mail_attachment":
                        exportedAttachment = ExportMailAttachment(command.ExportMailAttachmentRequest);
                        if (exportedAttachment is not null) _mailStore.UpsertExportedAttachment(exportedAttachment);
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
                        calendar = FilterCalendar(command.CalendarRequest);
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
                        mail = UpdateMailProperties(command.MailPropertiesRequest);
                        categories = new List<OutlookCategoryDto>(_mockCategories);
                        if (mail is not null) _mailStore.UpsertMail(mail);
                        _mailStore.SetCategories(categories);
                        break;
                    case "create_folder":
                        CreateFolder(command.CreateFolderRequest);
                        folderBatch = BuildFolderChildrenBatch(new FolderDiscoveryRequest
                        {
                            StoreId = FindFolder(command.CreateFolderRequest?.ParentFolderPath ?? string.Empty)?.StoreId ?? string.Empty,
                            ParentFolderPath = command.CreateFolderRequest?.ParentFolderPath ?? string.Empty,
                        });
                        _mailStore.ApplyFolderBatch(folderBatch);
                        break;
                    case "delete_folder":
                        var deletedParentPath = FindFolder(command.DeleteFolderRequest?.FolderPath ?? string.Empty)?.ParentFolderPath ?? string.Empty;
                        var deletedStoreId = FindFolder(command.DeleteFolderRequest?.FolderPath ?? string.Empty)?.StoreId ?? string.Empty;
                        DeleteFolder(command.DeleteFolderRequest);
                        folderBatch = BuildFolderChildrenBatch(new FolderDiscoveryRequest
                        {
                            StoreId = deletedStoreId,
                            ParentFolderPath = deletedParentPath,
                        });
                        _mailStore.ApplyFolderBatch(folderBatch);
                        break;
                    case "move_mail":
                        MoveMail(command.MoveMailRequest);
                        mails = SyncVisibleMails(command.MoveMailRequest?.MailId);
                        _mailStore.SetMails(mails);
                        folderBatch = BuildFolderCountsBatch(command.MoveMailRequest?.SourceFolderPath, command.MoveMailRequest?.DestinationFolderPath);
                        _mailStore.ApplyFolderBatch(folderBatch);
                        break;
                    case "delete_mail":
                        DeleteMail(command.DeleteMailRequest);
                        mails = SyncVisibleMails(command.DeleteMailRequest?.MailId);
                        _mailStore.SetMails(mails);
                        folderBatch = BuildFolderCountsBatch(command.DeleteMailRequest?.FolderPath);
                        _mailStore.ApplyFolderBatch(folderBatch);
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

            if (folderBatch is not null)
            {
                await _notifications.Clients.All.SendAsync("FolderSyncStarted", new FolderSyncBeginDto
                {
                    SyncId = folderBatch.SyncId,
                    Reset = true,
                    Timestamp = DateTime.Now,
                }, ct);
                await _notifications.Clients.All.SendAsync("FoldersPatched", folderBatch, ct);
                await _notifications.Clients.All.SendAsync("FolderSyncCompleted", new FolderSyncCompleteDto
                {
                    SyncId = folderBatch.SyncId,
                    TotalCount = folderBatch.Folders.Count,
                    Message = "Mock folder sync completed",
                    Timestamp = DateTime.Now,
                }, ct);
            }
            if (mails is not null) await _notifications.Clients.All.SendAsync("MailsUpdated", mails, ct);
            if (mailSearchSliceResult is not null)
            {
                await _notifications.Clients.All.SendAsync("MailSearchStarted", new MailSearchSliceResultDto
                {
                    SearchId = mailSearchSliceResult.SearchId,
                    Reset = mailSearchSliceResult.Reset,
                    Sequence = mailSearchSliceResult.Sequence,
                }, ct);
                await _notifications.Clients.All.SendAsync("MailSearchPatched", mailSearchSliceResult, ct);
            }
            if (mailSearchComplete is not null) await _notifications.Clients.All.SendAsync("MailSearchCompleted", mailSearchComplete, ct);
            if (mail is not null) await _notifications.Clients.All.SendAsync("MailUpdated", mail, ct);
            if (mailBody is not null) await _notifications.Clients.All.SendAsync("MailBodyUpdated", mailBody, ct);
            if (mailAttachments is not null) await _notifications.Clients.All.SendAsync("MailAttachmentsUpdated", mailAttachments, ct);
            if (exportedAttachment is not null) await _notifications.Clients.All.SendAsync("MailAttachmentExported", exportedAttachment, ct);
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

            var data = MockOutlookSeedFactory.Create();
            _mockStores = data.Stores;
            _mockFolders = data.Folders;
            _mockMails.AddRange(data.Mails);
            _mockCategories = data.Categories;
            _mockRules = data.Rules;
            _mockCalendar = data.Calendar;
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

        private MailBodyDto? FetchMailBody(FetchMailBodyRequest? request)
        {
            if (request is null || string.IsNullOrWhiteSpace(request.MailId)) return null;
            var mail = _mockMails.FirstOrDefault(item =>
                item.Id == request.MailId
                && (string.IsNullOrWhiteSpace(request.FolderPath) || item.FolderPath == request.FolderPath));
            if (mail is null) return null;
            return new MailBodyDto
            {
                MailId = mail.Id,
                FolderPath = mail.FolderPath,
                Body = mail.Body,
                BodyHtml = mail.BodyHtml,
            };
        }

        private MailAttachmentsDto? FetchMailAttachments(FetchMailAttachmentsRequest? request)
        {
            if (request is null || string.IsNullOrWhiteSpace(request.MailId)) return null;
            var mail = _mockMails.FirstOrDefault(item =>
                item.Id == request.MailId
                && (string.IsNullOrWhiteSpace(request.FolderPath) || item.FolderPath == request.FolderPath));
            if (mail is null) return null;

            var attachments = MockOutlookAttachmentFactory.Build(mail).Select(attachment =>
            {
                attachment.IsExported = false;
                return attachment;
            }).ToList();

            return new MailAttachmentsDto
            {
                MailId = mail.Id,
                FolderPath = mail.FolderPath,
                Attachments = attachments,
            };
        }

        private ExportedMailAttachmentDto? ExportMailAttachment(ExportMailAttachmentRequest? request)
        {
            if (request is null || string.IsNullOrWhiteSpace(request.MailId) || string.IsNullOrWhiteSpace(request.AttachmentId)) return null;
            var mail = _mockMails.FirstOrDefault(item =>
                item.Id == request.MailId
                && (string.IsNullOrWhiteSpace(request.FolderPath) || item.FolderPath == request.FolderPath));
            if (mail is null) return null;

            var attachment = MockOutlookAttachmentFactory.Build(mail).FirstOrDefault(item => item.AttachmentId == request.AttachmentId);
            if (attachment is null) return null;

            var exportPath = _attachmentExports.CreateExportPath(mail.Id, mail.Subject, mail.ReceivedTime, attachment.Name);
            File.WriteAllText(exportPath, $"Mock attachment exported from {mail.Subject}\n\nAttachment: {attachment.Name}\n", System.Text.Encoding.UTF8);
            var exported = new ExportedMailAttachmentDto
            {
                MailId = mail.Id,
                FolderPath = mail.FolderPath,
                AttachmentId = attachment.AttachmentId,
                ExportedAttachmentId = Guid.NewGuid().ToString(),
                Name = attachment.Name,
                ContentType = attachment.ContentType,
                Size = new FileInfo(exportPath).Length,
                ExportedPath = exportPath,
                ExportedAt = DateTime.Now,
            };
            return exported;
        }

        private List<CalendarEventDto> FilterCalendar(FetchCalendarRequest? request)
        {
            var start = request?.StartDate?.Date ?? DateTime.Now.Date;
            var end = request?.EndDate?.Date ?? start.AddDays(Math.Max(1, request?.DaysForward ?? 31));
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
                    Color = string.IsNullOrWhiteSpace(request.Color) ? "olCategoryColorNone" : request.Color,
                    ColorValue = request.ColorValue,
                    ShortcutKey = request.ShortcutKey,
                });
                return;
            }

            existing.Color = string.IsNullOrWhiteSpace(request.Color) ? existing.Color : request.Color;
            existing.ColorValue = request.ColorValue;
            existing.ShortcutKey = request.ShortcutKey;
        }

        private void UpdateMailMarker(MailMarkerCommandRequest? request, Action<MailItemDto> update)
        {
            if (request is null || string.IsNullOrWhiteSpace(request.MailId)) return;
            var mail = _mockMails.FirstOrDefault(item => item.Id == request.MailId);
            if (mail is null) return;
            update(mail);
        }

        private MailItemDto? UpdateMailProperties(MailPropertiesCommandRequest? request)
        {
            if (request is null) return null;
            var mail = _mockMails.FirstOrDefault(item => item.Id == request.MailId);
            if (mail is null) return null;

            if (request.IsRead.HasValue) mail.IsRead = request.IsRead.Value;
            ApplyFlag(mail, request);
            mail.Categories = string.Join(",", request.Categories.Where(category => !string.IsNullOrWhiteSpace(category)).Select(category => category.Trim()));

            foreach (var category in request.NewCategories)
                UpsertCategory(new CategoryCommandRequest { Name = category.Name, Color = category.Color, ColorValue = category.ColorValue, ShortcutKey = category.ShortcutKey });

            return CloneMail(mail);
        }

        private void CreateFolder(CreateFolderRequest? request)
        {
            if (request is null || string.IsNullOrWhiteSpace(request.ParentFolderPath) || string.IsNullOrWhiteSpace(request.Name)) return;
            var parent = FindFolder(request.ParentFolderPath);
            if (parent is null) return;
            var name = request.Name.Trim();
            if (_mockFolders.Any(folder => folder.ParentFolderPath == parent.FolderPath && folder.Name.Equals(name, StringComparison.OrdinalIgnoreCase))) return;
            var path = $"{request.ParentFolderPath}\\{name}";
            AddMockFolder(name, path, parent.FolderPath, parent.StoreId);
        }

        private void DeleteFolder(DeleteFolderRequest? request)
        {
            if (request is null || string.IsNullOrWhiteSpace(request.FolderPath)) return;
            DeleteFolderFrom(request.FolderPath);
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

        private void DeleteMail(DeleteMailRequest? request)
        {
            if (request is null || string.IsNullOrWhiteSpace(request.MailId)) return;
            MoveMail(new MoveMailRequest
            {
                MailId = request.MailId,
                SourceFolderPath = request.FolderPath,
                DestinationFolderPath = MockOutlookPaths.Deleted,
            });
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

        private static string BuildChatReply(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "Mock Outlook 收到空白訊息。";

            return $"Mock Outlook 已收到：「{text.Trim()}」。目前本機資料可直接測試 request、更新與即時廣播。";
        }

        private void RefreshFolderCounts()
        {
            foreach (var folder in _mockFolders)
            {
                folder.ItemCount = _mockMails.Count(mail => mail.FolderPath == folder.FolderPath);
                folder.HasChildren = _mockFolders.Any(child => child.ParentFolderPath == folder.FolderPath);
                folder.ChildrenLoaded = !folder.HasChildren;
                folder.DiscoveryState = folder.HasChildren ? "partial" : "loaded";
            }
        }

        private void AddMockFolder(string name, string folderPath, string parentFolderPath, string storeId, bool isStoreRoot = false)
        {
            _mockFolders.Add(new FolderDto
            {
                Name = name,
                EntryId = MockFolderEntryId(storeId, folderPath),
                FolderPath = folderPath,
                ParentEntryId = string.IsNullOrWhiteSpace(parentFolderPath) ? string.Empty : MockFolderEntryId(storeId, parentFolderPath),
                ParentFolderPath = parentFolderPath,
                StoreId = storeId,
                IsStoreRoot = isStoreRoot,
                DiscoveryState = "loaded",
            });
            RefreshFolderCounts();
        }

        private FolderDto? FindFolder(string path)
        {
            return _mockFolders.FirstOrDefault(folder => folder.FolderPath == path);
        }

        private void DeleteFolderFrom(string path)
        {
            _mockFolders.RemoveAll(folder =>
                folder.FolderPath == path
                || folder.ParentFolderPath.StartsWith(path, StringComparison.OrdinalIgnoreCase)
                || folder.FolderPath.StartsWith($"{path}\\", StringComparison.OrdinalIgnoreCase));
        }

        private static List<FolderDto> CloneFolders(List<FolderDto> folders)
        {
            return folders.Select(folder => new FolderDto
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
            }).ToList();
        }

        private FolderSyncBatchDto BuildFolderRootsBatch(bool reset)
        {
            RefreshFolderCounts();
            return new FolderSyncBatchDto
            {
                SyncId = Guid.NewGuid().ToString(),
                Sequence = 1,
                Reset = reset,
                IsFinal = true,
                Stores = _mockStores.Select(store => new OutlookStoreDto
                {
                    StoreId = store.StoreId,
                    DisplayName = store.DisplayName,
                    StoreKind = store.StoreKind,
                    StoreFilePath = store.StoreFilePath,
                    RootFolderPath = store.RootFolderPath,
                }).ToList(),
                Folders = CloneFolders(_mockFolders.Where(folder => folder.IsStoreRoot).ToList()),
            };
        }

        private FolderSyncBatchDto BuildFolderChildrenBatch(FolderDiscoveryRequest? request)
        {
            RefreshFolderCounts();
            request ??= new FolderDiscoveryRequest();
            var parent = _mockFolders.FirstOrDefault(folder =>
                string.Equals(folder.StoreId, request.StoreId, StringComparison.OrdinalIgnoreCase)
                && (
                    (!string.IsNullOrWhiteSpace(request.ParentEntryId) && string.Equals(folder.EntryId, request.ParentEntryId, StringComparison.OrdinalIgnoreCase))
                    || (!string.IsNullOrWhiteSpace(request.ParentFolderPath) && string.Equals(folder.FolderPath, request.ParentFolderPath, StringComparison.OrdinalIgnoreCase))
                ));

            var folders = new List<FolderDto>();
            if (parent is not null)
            {
                var loadedParent = new FolderDto
                {
                    Name = parent.Name,
                    EntryId = parent.EntryId,
                    FolderPath = parent.FolderPath,
                    ParentEntryId = parent.ParentEntryId,
                    ParentFolderPath = parent.ParentFolderPath,
                    ItemCount = parent.ItemCount,
                    StoreId = parent.StoreId,
                    IsStoreRoot = parent.IsStoreRoot,
                    HasChildren = parent.HasChildren,
                    ChildrenLoaded = true,
                    DiscoveryState = "loaded",
                };
                folders.Add(loadedParent);
                folders.AddRange(CloneFolders(_mockFolders
                    .Where(folder => string.Equals(folder.ParentFolderPath, parent.FolderPath, StringComparison.OrdinalIgnoreCase))
                    .Take(Math.Clamp(request.MaxChildren <= 0 ? 50 : request.MaxChildren, 1, 200))
                    .ToList()));
            }

            return new FolderSyncBatchDto
            {
                SyncId = string.IsNullOrWhiteSpace(request.SyncId) ? Guid.NewGuid().ToString() : request.SyncId,
                Sequence = 1,
                Reset = false,
                IsFinal = true,
                Stores = new List<OutlookStoreDto>(),
                Folders = folders,
            };
        }

        private FolderSyncBatchDto BuildFolderCountsBatch(params string?[] folderPaths)
        {
            RefreshFolderCounts();
            var requested = folderPaths
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Select(path => path!)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            return new FolderSyncBatchDto
            {
                SyncId = Guid.NewGuid().ToString(),
                Sequence = 1,
                Reset = false,
                IsFinal = true,
                Stores = new List<OutlookStoreDto>(),
                Folders = CloneFolders(_mockFolders.Where(folder => requested.Contains(folder.FolderPath)).ToList()),
            };
        }

        private static string MockFolderEntryId(string storeId, string folderPath)
        {
            return $"{storeId}:{folderPath}";
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

    }
}
