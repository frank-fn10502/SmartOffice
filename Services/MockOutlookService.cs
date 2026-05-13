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
                _mailStore.SetMails(MockOutlookMailSearch.FilterMails(_mockMails, MockOutlookPaths.Inbox, MockOutlookPaths.Inbox, 30, DateTime.Now.AddDays(-7)));
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
            List<MailSearchSliceResultDto>? mailSearchSliceResults = null;
            List<FolderMailsSliceResultDto>? folderMailsSliceResults = null;
            MailSearchCompleteDto? mailSearchComplete = null;
            FolderMailsCompleteDto? folderMailsComplete = null;
            List<MailItemDto>? mails = null;
            MailItemDto? mail = null;
            MailBodyDto? mailBody = null;
            MailAttachmentsDto? mailAttachments = null;
            MailConversationDto? mailConversation = null;
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
                        folderBatch = ToFullFolderBatch(folderBatch, _mailStore.ApplyFolderBatch(folderBatch));
                        break;
                    case "fetch_folder_children":
                        folderBatch = BuildFolderChildrenBatch(command.FolderDiscoveryRequest);
                        folderBatch = ToFullFolderBatch(folderBatch, _mailStore.ApplyFolderBatch(folderBatch));
                        break;
                    case "fetch_mails":
                        mails = MockOutlookMailSearch.FilterMails(
                            _mockMails,
                            MockOutlookPaths.Inbox,
                            command.MailsRequest?.FolderPath ?? MockOutlookPaths.Inbox,
                            command.MailsRequest?.MaxCount ?? 30,
                            command.MailsRequest?.ReceivedFrom,
                            command.MailsRequest?.ReceivedTo);
                        _mailStore.SetMails(mails);
                        break;
                    case "fetch_mail_search_slice":
                        var request = command.MailSearchSliceRequest ?? new MailSearchSliceRequest();
                        var searchResults = MockOutlookMailSearch.FetchMailSearchSlice(_mockMails, request);
                        mailSearchSliceResults = BuildMailSearchResultBatches(command.Id, request, searchResults);
                        _mailStore.BeginMailSearch(request.ResetSearchResults, request.SearchId);
                        foreach (var batch in mailSearchSliceResults)
                            _mailStore.ApplyMailSearchSliceResult(batch);
                        if (request.CompleteSearchOnSlice)
                        {
                            mailSearchComplete = new MailSearchCompleteDto
                            {
                                SearchId = request.SearchId,
                                CommandId = command.Id,
                                ParentCommandId = request.ParentCommandId,
                                TotalCount = _mailStore.GetMailSearchResultCount(request.SearchId),
                                Message = "Mock mail search completed",
                                Timestamp = DateTime.Now,
                            };
                        }
                        break;
                    case "fetch_folder_mails_slice":
                        var folderMailsRequest = command.FolderMailsSliceRequest ?? new FolderMailsSliceRequest();
                        var folderMailsResults = MockOutlookMailSearch.FetchFolderMailsSlice(_mockMails, folderMailsRequest);
                        folderMailsSliceResults = BuildFolderMailsResultBatches(command.Id, folderMailsRequest, folderMailsResults);
                        _mailStore.BeginFolderMails(folderMailsRequest.ResetResults, folderMailsRequest.FolderMailsId);
                        foreach (var batch in folderMailsSliceResults)
                            _mailStore.ApplyFolderMailsSliceResult(batch);
                        if (folderMailsRequest.CompleteOnSlice)
                        {
                            folderMailsComplete = new FolderMailsCompleteDto
                            {
                                FolderMailsId = folderMailsRequest.FolderMailsId,
                                CommandId = command.Id,
                                ParentCommandId = folderMailsRequest.ParentCommandId,
                                TotalCount = _mailStore.GetFolderMailResultCount(folderMailsRequest.FolderMailsId),
                                Message = "Mock folder mails completed",
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
                    case "fetch_mail_conversation":
                        mailConversation = FetchMailConversation(command.MailConversationRequest);
                        if (mailConversation is not null) _mailStore.SetMailConversation(mailConversation);
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
                    case "manage_rule":
                        ManageRule(command.RuleRequest);
                        rules = new List<OutlookRuleDto>(_mockRules);
                        _mailStore.SetRules(rules);
                        break;
                    case "fetch_calendar":
                        calendar = FilterCalendar(command.CalendarRequest);
                        _mailStore.SetCalendarEvents(calendar);
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
                        folderBatch = ToFullFolderBatch(folderBatch, _mailStore.ApplyFolderBatch(folderBatch));
                        break;
                    case "delete_folder":
                        DeleteFolder(command.DeleteFolderRequest);
                        folderBatch = BuildFullFolderBatch();
                        _mailStore.ApplyFolderBatch(folderBatch);
                        break;
                    case "move_mail":
                        MoveMail(command.MoveMailRequest);
                        mails = SyncVisibleMails(command.MoveMailRequest?.MailId);
                        _mailStore.SetMails(mails);
                        folderBatch = BuildFolderCountsBatch(command.MoveMailRequest?.SourceFolderPath, command.MoveMailRequest?.DestinationFolderPath);
                        folderBatch = ToFullFolderBatch(folderBatch, _mailStore.ApplyFolderBatch(folderBatch));
                        break;
                    case "move_mails":
                        MoveMails(command.MoveMailsRequest);
                        mails = SyncVisibleMails(command.MoveMailsRequest?.MailIds.FirstOrDefault());
                        _mailStore.SetMails(mails);
                        folderBatch = BuildFolderCountsBatch(MoveMailsCountPaths(command.MoveMailsRequest).ToArray());
                        folderBatch = ToFullFolderBatch(folderBatch, _mailStore.ApplyFolderBatch(folderBatch));
                        break;
                    case "delete_mail":
                        DeleteMail(command.DeleteMailRequest);
                        mails = SyncVisibleMails(command.DeleteMailRequest?.MailId);
                        _mailStore.SetMails(mails);
                        folderBatch = BuildFolderCountsBatch(command.DeleteMailRequest?.FolderPath);
                        folderBatch = ToFullFolderBatch(folderBatch, _mailStore.ApplyFolderBatch(folderBatch));
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
            if (mailSearchSliceResults is not null)
            {
                var firstBatch = mailSearchSliceResults.FirstOrDefault();
                await _notifications.Clients.All.SendAsync("MailSearchStarted", new MailSearchSliceResultDto
                {
                    SearchId = firstBatch?.SearchId ?? string.Empty,
                    Reset = firstBatch?.Reset ?? false,
                    Sequence = firstBatch?.Sequence ?? 0,
                }, ct);
                foreach (var batch in mailSearchSliceResults)
                    await _notifications.Clients.All.SendAsync("MailSearchPatched", batch, ct);
            }
            if (mailSearchComplete is not null) await _notifications.Clients.All.SendAsync("MailSearchCompleted", mailSearchComplete, ct);
            if (folderMailsSliceResults is not null)
            {
                var firstBatch = folderMailsSliceResults.FirstOrDefault();
                await _notifications.Clients.All.SendAsync("FolderMailsStarted", new FolderMailsSliceResultDto
                {
                    FolderMailsId = firstBatch?.FolderMailsId ?? string.Empty,
                    Reset = firstBatch?.Reset ?? false,
                    Sequence = firstBatch?.Sequence ?? 0,
                }, ct);
                foreach (var batch in folderMailsSliceResults)
                    await _notifications.Clients.All.SendAsync("FolderMailsPatched", batch, ct);
            }
            if (folderMailsComplete is not null) await _notifications.Clients.All.SendAsync("FolderMailsCompleted", folderMailsComplete, ct);
            if (mail is not null) await _notifications.Clients.All.SendAsync("MailUpdated", mail, ct);
            if (mailBody is not null) await _notifications.Clients.All.SendAsync("MailBodyUpdated", mailBody, ct);
            if (mailAttachments is not null) await _notifications.Clients.All.SendAsync("MailAttachmentsUpdated", mailAttachments, ct);
            if (mailConversation is not null) await _notifications.Clients.All.SendAsync("MailConversationUpdated", mailConversation, ct);
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

        private static List<MailSearchSliceResultDto> BuildMailSearchResultBatches(
            string commandId,
            MailSearchSliceRequest request,
            List<MailItemDto> searchResults)
        {
            var batchSize = Math.Clamp(request.ResultBatchSize <= 0 ? 5 : request.ResultBatchSize, 3, 5);
            var chunks = searchResults
                .Chunk(batchSize)
                .Select(chunk => chunk.ToList())
                .ToList();
            if (chunks.Count == 0) chunks.Add(new List<MailItemDto>());

            return chunks.Select((chunk, index) => new MailSearchSliceResultDto
            {
                SearchId = request.SearchId,
                CommandId = commandId,
                ParentCommandId = request.ParentCommandId,
                Sequence = (request.SliceIndex * 100000) + index + 1,
                SliceIndex = request.SliceIndex,
                SliceCount = request.SliceCount,
                Reset = request.ResetSearchResults && index == 0,
                IsFinal = request.CompleteSearchOnSlice && index == chunks.Count - 1,
                IsSliceComplete = index == chunks.Count - 1,
                Mails = chunk,
                Message = index == chunks.Count - 1 ? "Mock mail search slice completed" : "Mock mail search batch",
            }).ToList();
        }

        private static List<FolderMailsSliceResultDto> BuildFolderMailsResultBatches(
            string commandId,
            FolderMailsSliceRequest request,
            List<MailItemDto> mails)
        {
            var batchSize = Math.Clamp(request.ResultBatchSize <= 0 ? 5 : request.ResultBatchSize, 3, 5);
            var chunks = mails
                .Chunk(batchSize)
                .Select(chunk => chunk.ToList())
                .ToList();
            if (chunks.Count == 0) chunks.Add(new List<MailItemDto>());

            return chunks.Select((chunk, index) => new FolderMailsSliceResultDto
            {
                FolderMailsId = request.FolderMailsId,
                CommandId = commandId,
                ParentCommandId = request.ParentCommandId,
                Sequence = (request.SliceIndex * 100000) + index + 1,
                SliceIndex = request.SliceIndex,
                SliceCount = request.SliceCount,
                Reset = request.ResetResults && index == 0,
                IsFinal = request.CompleteOnSlice && index == chunks.Count - 1,
                IsSliceComplete = index == chunks.Count - 1,
                Mails = chunk,
                Message = index == chunks.Count - 1 ? "Mock folder mails slice completed" : "Mock folder mails batch",
            }).ToList();
        }

        private void ManageRule(OutlookRuleCommandRequest? request)
        {
            if (request is null) return;
            var originalName = string.IsNullOrWhiteSpace(request.OriginalRuleName)
                ? request.RuleName
                : request.OriginalRuleName;
            var index = _mockRules.FindIndex(rule =>
                string.Equals(rule.Name, originalName, StringComparison.OrdinalIgnoreCase)
                && (
                    request.OriginalExecutionOrder is null
                    || rule.ExecutionOrder == request.OriginalExecutionOrder.Value
                ));

            if (string.Equals(request.Operation, "delete", StringComparison.OrdinalIgnoreCase))
            {
                if (index >= 0) _mockRules.RemoveAt(index);
                ReorderMockRules();
                return;
            }

            if (string.Equals(request.Operation, "set_enabled", StringComparison.OrdinalIgnoreCase))
            {
                if (index >= 0) _mockRules[index].Enabled = request.Enabled;
                return;
            }

            var next = BuildMockRule(request);
            if (index >= 0) _mockRules[index] = next;
            else _mockRules.Insert(0, next);
            ReorderMockRules();
        }

        private static OutlookRuleDto BuildMockRule(OutlookRuleCommandRequest request)
        {
            return new OutlookRuleDto
            {
                StoreId = request.StoreId,
                Name = request.RuleName,
                Enabled = request.Enabled,
                ExecutionOrder = request.ExecutionOrder ?? 1,
                RuleType = request.RuleType,
                CanModifyDefinition = true,
                Conditions = BuildRuleConditionSummaries(request.Conditions),
                Actions = BuildRuleActionSummaries(request.Actions),
                Exceptions = new List<string>(),
            };
        }

        private static List<string> BuildRuleConditionSummaries(OutlookRuleConditionsRequest conditions)
        {
            var result = new List<string>();
            if (conditions.SubjectContains.Count > 0) result.Add($"Subject: Text={string.Join(", ", conditions.SubjectContains)}");
            if (conditions.BodyContains.Count > 0) result.Add($"Body: Text={string.Join(", ", conditions.BodyContains)}");
            if (conditions.SenderAddressContains.Count > 0) result.Add($"SenderAddress: Address={string.Join(", ", conditions.SenderAddressContains)}");
            if (conditions.Categories.Count > 0) result.Add($"Category: Categories={string.Join(", ", conditions.Categories)}");
            if (conditions.HasAttachment == true) result.Add("HasAttachment: (enabled)");
            return result;
        }

        private static List<string> BuildRuleActionSummaries(OutlookRuleActionsRequest actions)
        {
            var result = new List<string>();
            if (!string.IsNullOrWhiteSpace(actions.MoveToFolderPath)) result.Add($"MoveToFolder: FolderPath={actions.MoveToFolderPath}");
            if (actions.AssignCategories.Count > 0) result.Add($"AssignToCategory: Categories={string.Join(", ", actions.AssignCategories)}");
            if (actions.MarkAsTask) result.Add("MarkAsTask: (enabled)");
            if (actions.StopProcessingMoreRules) result.Add("Stop: (enabled)");
            return result;
        }

        private void ReorderMockRules()
        {
            _mockRules = _mockRules
                .OrderBy(rule => rule.ExecutionOrder <= 0 ? int.MaxValue : rule.ExecutionOrder)
                .ThenBy(rule => rule.Name)
                .Select((rule, index) =>
                {
                    rule.ExecutionOrder = index + 1;
                    return rule;
                })
                .ToList();
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

        private MailConversationDto? FetchMailConversation(FetchMailConversationRequest? request)
        {
            if (request is null || string.IsNullOrWhiteSpace(request.MailId)) return null;
            var mail = _mockMails.FirstOrDefault(item =>
                item.Id == request.MailId
                && (string.IsNullOrWhiteSpace(request.FolderPath) || item.FolderPath == request.FolderPath));
            if (mail is null) return null;

            var maxCount = Math.Clamp(request.MaxCount <= 0 ? 100 : request.MaxCount, 1, 300);
            var conversationId = string.IsNullOrWhiteSpace(mail.ConversationId)
                ? MockConversationId(mail.Subject)
                : mail.ConversationId;
            var topic = string.IsNullOrWhiteSpace(mail.ConversationTopic)
                ? NormalizeConversationTopic(mail.Subject)
                : mail.ConversationTopic;
            var mails = _mockMails
                .Where(item =>
                    (!string.IsNullOrWhiteSpace(item.ConversationId) && item.ConversationId == conversationId)
                    || string.Equals(NormalizeConversationTopic(item.Subject), topic, StringComparison.OrdinalIgnoreCase))
                .OrderBy(item => item.ReceivedTime)
                .Take(maxCount)
                .Select(CloneMail)
                .ToList();

            if (!request.IncludeBody)
            {
                foreach (var item in mails)
                {
                    item.Body = string.Empty;
                    item.BodyHtml = string.Empty;
                }
            }

            return new MailConversationDto
            {
                MailId = mail.Id,
                FolderPath = mail.FolderPath,
                ConversationId = conversationId,
                ConversationTopic = topic,
                Mails = mails,
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
            MoveFolderToDeletedItems(request.FolderPath);
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

        private void MoveMails(MoveMailsRequest? request)
        {
            if (request is null || request.MailIds.Count == 0 || string.IsNullOrWhiteSpace(request.DestinationFolderPath)) return;
            var mailIds = request.MailIds
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (var mail in _mockMails.Where(item => mailIds.Contains(item.Id)))
            {
                if (mail.FolderPath == request.DestinationFolderPath) continue;
                mail.FolderPath = request.DestinationFolderPath;
            }
            RefreshFolderCounts();
        }

        private static IEnumerable<string?> MoveMailsCountPaths(MoveMailsRequest? request)
        {
            if (request is null) yield break;
            yield return request.SourceFolderPath;
            foreach (var path in request.SourceFolderPaths) yield return path;
            yield return request.DestinationFolderPath;
        }

        private void DeleteMail(DeleteMailRequest? request)
        {
            if (request is null || string.IsNullOrWhiteSpace(request.MailId)) return;
            var mail = _mockMails.FirstOrDefault(item => item.Id == request.MailId);
            if (mail is null || IsInDefaultDeletedItems(mail.FolderPath)) return;
            var deletedFolder = DefaultDeletedFolderFor(mail.FolderPath);
            if (deletedFolder is null) return;

            MoveMail(new MoveMailRequest
            {
                MailId = request.MailId,
                SourceFolderPath = request.FolderPath,
                DestinationFolderPath = deletedFolder.FolderPath,
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

        private static string MockConversationId(string subject)
        {
            return $"mock-conv-{NormalizeConversationTopic(subject).ToLowerInvariant().Replace(" ", "-")}";
        }

        private static string NormalizeConversationTopic(string subject)
        {
            var topic = subject.Trim();
            while (true)
            {
                var normalized = topic.TrimStart();
                if (normalized.StartsWith("Re:", StringComparison.OrdinalIgnoreCase))
                {
                    topic = normalized[3..].TrimStart();
                    continue;
                }
                if (normalized.StartsWith("FW:", StringComparison.OrdinalIgnoreCase))
                {
                    topic = normalized[3..].TrimStart();
                    continue;
                }
                if (normalized.StartsWith("Fwd:", StringComparison.OrdinalIgnoreCase))
                {
                    topic = normalized[4..].TrimStart();
                    continue;
                }
                return normalized;
            }
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
                FolderType = isStoreRoot ? OutlookFolderType.StoreRoot : OutlookFolderType.Mail,
                DefaultItemType = isStoreRoot ? -1 : 0,
                DiscoveryState = "loaded",
            });
            RefreshFolderCounts();
        }

        private FolderDto? FindFolder(string path)
        {
            return _mockFolders.FirstOrDefault(folder => folder.FolderPath == path);
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
                FolderType = folder.FolderType,
                DefaultItemType = folder.DefaultItemType,
                IsHidden = folder.IsHidden,
                IsSystem = folder.IsSystem,
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

        private FolderSyncBatchDto BuildFullFolderBatch()
        {
            RefreshFolderCounts();
            return new FolderSyncBatchDto
            {
                SyncId = Guid.NewGuid().ToString(),
                Sequence = 1,
                Reset = true,
                IsFinal = true,
                Stores = _mockStores.Select(store => new OutlookStoreDto
                {
                    StoreId = store.StoreId,
                    DisplayName = store.DisplayName,
                    StoreKind = store.StoreKind,
                    StoreFilePath = store.StoreFilePath,
                    RootFolderPath = store.RootFolderPath,
                }).ToList(),
                Folders = CloneFolders(_mockFolders),
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
                    FolderType = parent.FolderType,
                    DefaultItemType = parent.DefaultItemType,
                    IsHidden = parent.IsHidden,
                    IsSystem = parent.IsSystem,
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

        private void MoveFolderToDeletedItems(string folderPath)
        {
            var target = FindFolder(folderPath);
            if (target is null || target.IsStoreRoot || target.IsHidden || target.IsSystem) return;

            var deletedFolder = DefaultDeletedFolderFor(target.FolderPath);

            if (deletedFolder is null) return;
            if (IsInDefaultDeletedItems(target.FolderPath)) return;

            var oldRootPath = target.FolderPath;
            var newRootPath = UniqueFolderPath(deletedFolder.FolderPath, target.Name, target.StoreId, oldRootPath);
            var movedFolders = _mockFolders
                .Where(folder =>
                    string.Equals(folder.FolderPath, oldRootPath, StringComparison.OrdinalIgnoreCase)
                    || folder.FolderPath.StartsWith($"{oldRootPath}\\", StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (var folder in movedFolders)
            {
                var oldFolderPath = folder.FolderPath;
                folder.FolderPath = RebasePath(oldFolderPath, oldRootPath, newRootPath);
                folder.ParentFolderPath = string.Equals(oldFolderPath, oldRootPath, StringComparison.OrdinalIgnoreCase)
                    ? deletedFolder.FolderPath
                    : RebasePath(folder.ParentFolderPath, oldRootPath, newRootPath);
                folder.EntryId = MockFolderEntryId(folder.StoreId, folder.FolderPath);
                folder.ParentEntryId = MockFolderEntryId(folder.StoreId, folder.ParentFolderPath);
            }

            foreach (var mail in _mockMails.Where(mail =>
                string.Equals(mail.FolderPath, oldRootPath, StringComparison.OrdinalIgnoreCase)
                || mail.FolderPath.StartsWith($"{oldRootPath}\\", StringComparison.OrdinalIgnoreCase)))
            {
                mail.FolderPath = RebasePath(mail.FolderPath, oldRootPath, newRootPath);
            }
        }

        private string UniqueFolderPath(string parentFolderPath, string folderName, string storeId, string currentPath)
        {
            var basePath = $"{parentFolderPath}\\{folderName}";
            var path = basePath;
            var index = 2;
            while (_mockFolders.Any(folder =>
                !string.Equals(folder.FolderPath, currentPath, StringComparison.OrdinalIgnoreCase)
                && string.Equals(folder.StoreId, storeId, StringComparison.OrdinalIgnoreCase)
                && string.Equals(folder.FolderPath, path, StringComparison.OrdinalIgnoreCase)))
            {
                path = $"{basePath} ({index})";
                index += 1;
            }

            return path;
        }

        private FolderDto? DefaultDeletedFolderFor(string folderPath)
        {
            var target = FindFolder(folderPath);
            var storeId = target?.StoreId ?? _mockFolders.FirstOrDefault(folder =>
                !folder.IsStoreRoot
                && folderPath.StartsWith($"{folder.FolderPath}\\", StringComparison.OrdinalIgnoreCase))?.StoreId;

            return _mockFolders.FirstOrDefault(folder =>
                folder.FolderType == OutlookFolderType.Deleted
                && (
                    string.IsNullOrWhiteSpace(storeId)
                    || string.Equals(folder.StoreId, storeId, StringComparison.OrdinalIgnoreCase)
                ));
        }

        private bool IsInDefaultDeletedItems(string folderPath)
        {
            var deletedFolder = DefaultDeletedFolderFor(folderPath);
            return deletedFolder is not null
                && (
                    string.Equals(folderPath, deletedFolder.FolderPath, StringComparison.OrdinalIgnoreCase)
                    || folderPath.StartsWith($"{deletedFolder.FolderPath}\\", StringComparison.OrdinalIgnoreCase)
                );
        }

        private static string RebasePath(string path, string oldRootPath, string newRootPath)
        {
            if (string.Equals(path, oldRootPath, StringComparison.OrdinalIgnoreCase)) return newRootPath;
            if (!path.StartsWith($"{oldRootPath}\\", StringComparison.OrdinalIgnoreCase)) return path;
            return $"{newRootPath}{path[oldRootPath.Length..]}";
        }

        private static FolderSyncBatchDto ToFullFolderBatch(FolderSyncBatchDto source, FolderSnapshotDto snapshot)
        {
            return new FolderSyncBatchDto
            {
                SyncId = source.SyncId,
                Sequence = source.Sequence,
                Reset = true,
                IsFinal = source.IsFinal,
                Stores = snapshot.Stores,
                Folders = snapshot.Folders,
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

        private static CalendarEventDto CloneCalendarEvent(CalendarEventDto item)
        {
            return new CalendarEventDto
            {
                Id = item.Id,
                Subject = item.Subject,
                Start = item.Start,
                End = item.End,
                Location = item.Location,
                Organizer = CloneRecipient(item.Organizer),
                RequiredAttendees = CloneRecipients(item.RequiredAttendees),
                IsRecurring = item.IsRecurring,
                BusyStatus = item.BusyStatus,
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

    }
}
