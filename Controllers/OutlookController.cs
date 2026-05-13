using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.SignalR;
using SmartOffice.Hub.Hubs;
using SmartOffice.Hub.Models;
using SmartOffice.Hub.Services;

namespace SmartOffice.Hub.Controllers
{
    [ApiController]
    [ApiExplorerSettings(GroupName = "outlook-v1")]
    [Route("api/outlook")]
    public class OutlookController : ControllerBase
    {
        private const int MaxMoveMailsBatchSize = 500;

        private readonly MailStore _mailStore;
        private readonly ChatStore _chatStore;
        private readonly CommandResultStore _commandResults;
        private readonly OutlookCommandQueue _commandQueue;
        private readonly MockOutlookService _mockOutlook;
        private readonly IHubContext<NotificationHub> _hub;
        private readonly AddinStatusStore _addinStatus;
        private readonly AttachmentExportService _attachmentExports;
        private readonly OutlookFolderCacheService _folderCache;

        public OutlookController(MailStore mailStore, ChatStore chatStore,
            CommandResultStore commandResults, OutlookCommandQueue commandQueue, MockOutlookService mockOutlook, IHubContext<NotificationHub> hub, AddinStatusStore addinStatus, AttachmentExportService attachmentExports, OutlookFolderCacheService folderCache)
        {
            _mailStore = mailStore;
            _chatStore = chatStore;
            _commandResults = commandResults;
            _commandQueue = commandQueue;
            _mockOutlook = mockOutlook;
            _hub = hub;
            _addinStatus = addinStatus;
            _attachmentExports = attachmentExports;
            _folderCache = folderCache;
        }

        // ===================== Web UI 呼叫這些 endpoint =====================

        /// <summary>
        /// Web UI、AI 或 MCP client 要求 mails；Hub 會透過 SignalR dispatch 給 Outlook AddIn。
        /// </summary>
        [HttpPost("request-mails")]
        public async Task<IActionResult> RequestMails([FromBody] RequestMailsApiRequest req, CancellationToken ct)
        {
            var normalized = BuildFetchMailsRequest(req);
            if (normalized.Error is not null) return normalized.Error;

            var cmd = new PendingCommand
            {
                Type = "fetch_mails",
                MailsRequest = normalized.Request
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 要求單封 mail body；mail list 本身只應先載入 metadata。
        /// </summary>
        [HttpPost("request-mail-body")]
        public async Task<IActionResult> RequestMailBody([FromBody] FetchMailBodyRequest req, CancellationToken ct)
        {
            var cmd = new PendingCommand
            {
                Type = "fetch_mail_body",
                MailBodyRequest = req
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 要求單封 mail attachment metadata。
        /// </summary>
        [HttpPost("request-mail-attachments")]
        public async Task<IActionResult> RequestMailAttachments([FromBody] FetchMailAttachmentsRequest req, CancellationToken ct)
        {
            var cmd = new PendingCommand
            {
                Type = "fetch_mail_attachments",
                MailAttachmentsRequest = req
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 要求指定 mail 所屬 Outlook conversation。
        /// </summary>
        [HttpPost("request-mail-conversation")]
        public async Task<IActionResult> RequestMailConversation([FromBody] FetchMailConversationRequest req, CancellationToken ct)
        {
            req ??= new FetchMailConversationRequest();
            req.MaxCount = Math.Clamp(req.MaxCount <= 0 ? 100 : req.MaxCount, 1, 300);
            var cmd = new PendingCommand
            {
                Type = "fetch_mail_conversation",
                MailConversationRequest = req
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 要求 AddIn 將指定 attachment 匯出到本機約定目錄。
        /// </summary>
        [HttpPost("request-export-mail-attachment")]
        public async Task<IActionResult> RequestExportMailAttachment([FromBody] ExportMailAttachmentRequest req, CancellationToken ct)
        {
            if (string.IsNullOrWhiteSpace(req.ExportRootPath))
                req.ExportRootPath = _attachmentExports.RootPath;

            var cmd = new PendingCommand
            {
                Type = "export_mail_attachment",
                ExportMailAttachmentRequest = req
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        /// <summary>
        /// Web UI Host 開啟已匯出的附件；只接受 Hub 已記錄的 exported attachment id。
        /// </summary>
        [HttpPost("open-exported-attachment")]
        public IActionResult OpenExportedAttachment([FromBody] OpenExportedAttachmentRequest req)
        {
            if (!_mailStore.TryGetExportedAttachment(req.ExportedAttachmentId, out var attachment))
                return NotFound(new { status = "not_found" });

            try
            {
                _attachmentExports.OpenExportedFile(attachment.ExportedPath);
                return Ok(new { status = "opened" });
            }
            catch (Exception ex)
            {
                return BadRequest(new { status = "open_failed", message = ex.Message });
            }
        }

        [HttpGet("attachment-export-settings")]
        public IActionResult GetAttachmentExportSettings()
        {
            return Ok(_attachmentExports.GetSettings());
        }

        [HttpPost("attachment-export-settings")]
        public IActionResult UpdateAttachmentExportSettings([FromBody] UpdateAttachmentExportSettingsRequest req)
        {
            try
            {
                return Ok(_attachmentExports.UpdateSettings(req.RootPath));
            }
            catch (Exception ex)
            {
                return BadRequest(new { status = "update_failed", message = ex.Message });
            }
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 要求 folder list。
        /// </summary>
        [HttpPost("request-folders")]
        public IActionResult RequestFolders(CancellationToken ct)
        {
            var command = new PendingCommand { Type = "fetch_folder_roots" };
            _commandResults.RecordDispatched(command);
            _ = Task.Run(() => _commandQueue.ExecuteExclusiveAsync(
                operationCt => RequestFoldersQueuedAsync(command, operationCt),
                CancellationToken.None));
            return Ok(OperationAccepted(command));
        }

        [HttpPost("request-folder-children")]
        public IActionResult RequestFolderChildren([FromBody] FolderDiscoveryRequest req, CancellationToken ct)
        {
            req ??= new FolderDiscoveryRequest();
            req.ParentFolderPath = OutlookFolderPathMapper.ToAddinPath(req.ParentFolderPath);
            req.MaxDepth = Math.Clamp(req.MaxDepth <= 0 ? 1 : req.MaxDepth, 1, 3);
            req.MaxChildren = Math.Clamp(req.MaxChildren <= 0 ? 50 : req.MaxChildren, 1, 200);
            req.Reset = false;

            var command = new PendingCommand
            {
                Type = "fetch_folder_children",
                FolderDiscoveryRequest = req,
            };
            _commandResults.RecordDispatched(command);
            _ = Task.Run(() => _commandQueue.ExecuteExclusiveAsync(
                operationCt => RequestFolderChildrenQueuedAsync(command, req, operationCt),
                CancellationToken.None));
            return Ok(OperationAccepted(command));
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 要求 Hub 封裝 folder tree discovery，並回傳符合名稱或 path 的 folder 候選。
        /// </summary>
        [HttpPost("request-find-folder")]
        public IActionResult RequestFindFolder([FromBody] FindFolderRequest req, CancellationToken ct)
        {
            req ??= new FindFolderRequest();
            req.FolderPath = OutlookFolderPathMapper.ToAddinPath(req.FolderPath);
            req.MaxResults = Math.Clamp(req.MaxResults <= 0 ? 20 : req.MaxResults, 1, 100);

            if (string.IsNullOrWhiteSpace(req.Name)
                && string.IsNullOrWhiteSpace(req.FolderPath)
                && string.IsNullOrWhiteSpace(req.FolderType))
            {
                return BadRequest(new
                {
                    status = "missing_folder_query",
                    message = "name, folderPath, or folderType is required."
                });
            }

            var command = new PendingCommand
            {
                Type = "find_folder",
                FindFolderRequest = req,
            };
            _commandResults.RecordDispatched(command);
            _ = Task.Run(() => _commandQueue.ExecuteExclusiveAsync(
                operationCt => RequestFindFolderQueuedAsync(command, operationCt),
                CancellationToken.None));
            return Ok(OperationAccepted(command, new
            {
                name = req.Name,
                folderPath = OutlookFolderPathMapper.ToApiPath(req.FolderPath),
                folderType = req.FolderType,
                storeId = req.StoreId,
                maxResults = req.MaxResults,
            }));
        }

        private async Task<IActionResult> RequestFolderChildrenQueuedAsync(PendingCommand command, FolderDiscoveryRequest req, CancellationToken ct)
        {
            var result = await _commandQueue.ExecuteQueuedCommandAsync(
                command,
                () => _mailStore.IsFolderChildrenLoaded(req.StoreId, req.ParentEntryId, req.ParentFolderPath),
                ensureReady: true,
                ct: ct);
            var body = ResultEnvelope(command.Id, command.Type, ResultState(result.Status), result.Message);
            return result.Success ? Ok(body) : StatusCode(result.HttpStatusCode, body);
        }

        private async Task<IActionResult> RequestFoldersQueuedAsync(PendingCommand command, CancellationToken ct)
        {
            var result = await _commandQueue.ExecuteQueuedCommandAsync(
                command,
                () => _mailStore.CountStoreRoots() > 0,
                ensureReady: true,
                ct: ct);
            var snapshot = _mailStore.GetFolderSnapshot();
            var body = new
            {
                requestId = command.Id,
                request = RequestName(command.Type),
                state = ResultState(result.Status),
                message = result.Message,
                data = new
                {
                    stores = snapshot.Stores.Count,
                    folders = snapshot.Folders.Count,
                },
            };
            return result.Success ? Ok(body) : StatusCode(result.HttpStatusCode, body);
        }

        private async Task<bool> RequestFindFolderQueuedAsync(PendingCommand command, CancellationToken ct)
        {
            var success = await _folderCache.EnsureFolderCacheAsync(command, ct, loadPendingChildren: true);
            _commandResults.RecordResult(new OutlookCommandResult
            {
                CommandId = command.Id,
                Success = success,
                Message = success ? "completed" : "folder_cache_unavailable",
                Timestamp = DateTime.Now,
            });
            return success;
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 要求 Outlook rule list。
        /// </summary>
        [HttpPost("request-rules")]
        public async Task<IActionResult> RequestRules(CancellationToken ct)
        {
            return await DispatchCommandAsync(new PendingCommand { Type = "fetch_rules" }, ct);
        }

        [HttpPost("request-manage-rule")]
        public async Task<IActionResult> RequestManageRule([FromBody] OutlookRuleCommandRequest req, CancellationToken ct)
        {
            var error = ValidateRuleRequest(req);
            if (error is not null) return error;

            var cmd = new PendingCommand
            {
                Type = "manage_rule",
                RuleRequest = req,
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 要求 Outlook master category list。
        /// </summary>
        [HttpPost("request-categories")]
        public async Task<IActionResult> RequestCategories(CancellationToken ct)
        {
            return await DispatchCommandAsync(new PendingCommand { Type = "fetch_categories" }, ct);
        }

        /// <summary>
        /// Web UI 或工作機測試用：透過正式 SignalR channel 發送 ping command。
        /// </summary>
        [HttpPost("request-signalr-ping")]
        public async Task<IActionResult> RequestSignalRPing(CancellationToken ct)
        {
            return await DispatchCommandAsync(new PendingCommand { Type = "ping" }, ct);
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 要求 Outlook calendar events。
        /// </summary>
        [HttpPost("request-calendar")]
        public async Task<IActionResult> RequestCalendar([FromBody] FetchCalendarRequest? req, CancellationToken ct)
        {
            var cmd = new PendingCommand
            {
                Type = "fetch_calendar",
                CalendarRequest = req ?? new FetchCalendarRequest()
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        [HttpPost("request-update-mail-properties")]
        public async Task<IActionResult> RequestUpdateMailProperties([FromBody] MailPropertiesCommandRequest req, CancellationToken ct)
        {
            if (!CanUseMailMutation(req?.MailId, out var unsupported))
                return UnsupportedMailMutation(unsupported, "update_mail_properties");

            var cmd = new PendingCommand
            {
                Type = "update_mail_properties",
                MailPropertiesRequest = req
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        [HttpPost("request-upsert-category")]
        public async Task<IActionResult> RequestUpsertCategory([FromBody] CategoryCommandRequest req, CancellationToken ct)
        {
            var cmd = new PendingCommand
            {
                Type = "upsert_category",
                CategoryRequest = req
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        [HttpPost("request-create-folder")]
        public async Task<IActionResult> RequestCreateFolder([FromBody] CreateFolderRequest req, CancellationToken ct)
        {
            var cmd = new PendingCommand
            {
                Type = "create_folder",
                CreateFolderRequest = req
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        [HttpPost("request-delete-folder")]
        public async Task<IActionResult> RequestDeleteFolder([FromBody] DeleteFolderRequest req, CancellationToken ct)
        {
            var cmd = new PendingCommand
            {
                Type = "delete_folder",
                DeleteFolderRequest = req
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        [HttpPost("request-move-mail")]
        public async Task<IActionResult> RequestMoveMail([FromBody] MoveMailRequest req, CancellationToken ct)
        {
            if (!CanUseMailMutation(req?.MailId, out var unsupported))
                return UnsupportedMailMutation(unsupported, "move_mail");

            var cmd = new PendingCommand
            {
                Type = "move_mail",
                MoveMailRequest = req
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        [HttpPost("request-move-mails")]
        public async Task<IActionResult> RequestMoveMails([FromBody] MoveMailsRequest req, CancellationToken ct)
        {
            req.MailIds ??= new List<string>();
            req.SourceFolderPaths ??= new List<string>();

            if (req.MailIds.Count > MaxMoveMailsBatchSize)
            {
                return BadRequest(new
                {
                    status = "too_many_mail_ids",
                    message = $"move_mails 單次最多只能移動 {MaxMoveMailsBatchSize} 封郵件；請由 caller 分批呼叫。",
                    maxBatchSize = MaxMoveMailsBatchSize,
                    actualCount = req.MailIds.Count
                });
            }

            var unsupportedMailIds = req.MailIds
                .Where(mailId => !CanUseMailMutation(mailId, out _))
                .ToList();
            if (unsupportedMailIds.Count > 0)
            {
                return BadRequest(new
                {
                    status = "unsupported_outlook_item_type",
                    operation = "move_mails",
                    message = "部分項目不是一般郵件，不能使用 mail mutation；會議邀請請使用 Outlook 會議/行事曆流程。",
                    mailIds = unsupportedMailIds
                });
            }

            var cmd = new PendingCommand
            {
                Type = "move_mails",
                MoveMailsRequest = req
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        [HttpPost("request-delete-mail")]
        public async Task<IActionResult> RequestDeleteMail([FromBody] DeleteMailRequest req, CancellationToken ct)
        {
            if (!CanUseMailMutation(req?.MailId, out var unsupported))
                return UnsupportedMailMutation(unsupported, "delete_mail");

            var cmd = new PendingCommand
            {
                Type = "delete_mail",
                DeleteMailRequest = req
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        private bool CanUseMailMutation(string? mailId, out MailItemDto? unsupported)
        {
            unsupported = null;
            if (string.IsNullOrWhiteSpace(mailId)) return true;

            var cached = _mailStore.FindCachedMail(mailId);
            if (cached is null) return true;

            var messageClass = cached.MessageClass?.Trim() ?? "";
            if (messageClass.Length == 0 || string.Equals(messageClass, "IPM.Note", StringComparison.OrdinalIgnoreCase))
                return true;

            unsupported = cached;
            return false;
        }

        private IActionResult UnsupportedMailMutation(MailItemDto? mail, string operation)
        {
            return BadRequest(new
            {
                status = "unsupported_outlook_item_type",
                operation,
                mailId = mail?.Id ?? "",
                messageClass = mail?.MessageClass ?? "",
                message = "此 Outlook item 不是一般郵件，不能使用 mail mutation；會議邀請請使用 Outlook 會議/行事曆流程。"
            });
        }

        private Task<IActionResult> DispatchCommandAsync(PendingCommand cmd, CancellationToken ct)
        {
            OutlookFolderPathMapper.NormalizeRequest(cmd);

            // Hub 已收下的 AddIn request 要由中心 queue 控制完成，避免外部 client 斷線造成 Outlook automation 半路取消。
            _commandResults.RecordDispatched(cmd);
            _ = Task.Run(() => _commandQueue.ExecuteExclusiveAsync(async operationCt =>
            {
                if (RequiresFolderCache(cmd) && !await _folderCache.EnsureFolderCacheAsync(cmd, operationCt))
                {
                    _commandResults.RecordResult(new OutlookCommandResult
                    {
                        CommandId = cmd.Id,
                        Success = false,
                        Message = "folder_cache_unavailable",
                        Timestamp = DateTime.Now,
                    });
                    return false;
                }

                if (IsBlockedFolderDelete(cmd))
                {
                    _commandResults.RecordResult(new OutlookCommandResult
                    {
                        CommandId = cmd.Id,
                        Success = false,
                        Message = "manual_delete_required",
                        Timestamp = DateTime.Now,
                    });
                    return false;
                }

                var result = await _commandQueue.ExecuteQueuedCommandAsync(cmd, DataReadyPredicate(cmd), ct: operationCt);
                await _hub.Clients.All.SendAsync("AddinStatus", _addinStatus.GetStatus(), CancellationToken.None);
                return true;
            }, CancellationToken.None));

            return Task.FromResult<IActionResult>(Ok(OperationAccepted(cmd)));
        }

        private static object OperationAccepted(PendingCommand command, object? data = null)
        {
            return ResultEnvelope(
                command.Id,
                command.Type,
                "accepted",
                "Request accepted. Poll the paired fetch-result-* endpoint for state and data.",
                data);
        }

        private static object ResultEnvelope(string requestId, string commandType, string state, string message, object? data = null)
        {
            return new
            {
                requestId,
                request = RequestName(commandType),
                state,
                message,
                data = data ?? new { },
            };
        }

        private (FetchMailsRequest Request, IActionResult? Error) BuildFetchMailsRequest(RequestMailsApiRequest req)
        {
            req ??= new RequestMailsApiRequest();
            var request = new FetchMailsRequest
            {
                FolderPath = req.FolderPath,
                ReceivedFrom = UtcDateTime.Normalize(req.ReceivedFrom),
                ReceivedTo = UtcDateTime.Normalize(req.ReceivedTo),
                MaxCount = req.MaxCount,
            };
            if (req.LookbackHours is null) return (request, null);

            if (req.LookbackHours <= 0)
            {
                return (request, BadRequest(new
                {
                    status = "invalid_lookback_hours",
                    message = "lookbackHours 必須大於 0，例如 12 代表過去 12 小時、24 代表過去 1 天。"
                }));
            }

            if (req.LookbackHours > 24 * 365)
            {
                return (request, BadRequest(new
                {
                    status = "invalid_lookback_hours",
                    message = "lookbackHours 不可超過 8760 小時。若需要更大的範圍，請改用 receivedFrom / receivedTo。"
                }));
            }

            request.ReceivedTo ??= UtcDateTime.Now;
            request.ReceivedFrom ??= request.ReceivedTo.Value.Subtract(TimeSpan.FromHours(req.LookbackHours.Value));
            request.ReceivedFrom = UtcDateTime.Normalize(request.ReceivedFrom);
            request.ReceivedTo = UtcDateTime.Normalize(request.ReceivedTo);
            return (request, null);
        }

        private bool IsBlockedFolderDelete(PendingCommand cmd)
        {
            if (!string.Equals(cmd.Type, "delete_folder", StringComparison.OrdinalIgnoreCase)) return false;
            var folderPath = cmd.DeleteFolderRequest?.FolderPath;
            return !string.IsNullOrWhiteSpace(folderPath) && IsInDefaultDeletedItems(folderPath);
        }

        private bool IsInDefaultDeletedItems(string folderPath)
        {
            var addinPath = OutlookFolderPathMapper.ToAddinPath(folderPath);
            if (string.IsNullOrWhiteSpace(addinPath)) return false;

            var folders = _mailStore.GetFolderSnapshot().Folders;
            var target = folders.FirstOrDefault(folder =>
                string.Equals(folder.FolderPath, addinPath, StringComparison.OrdinalIgnoreCase));
            var storeId = target?.StoreId ?? folders.FirstOrDefault(folder =>
                !folder.IsStoreRoot
                && addinPath.StartsWith($"{folder.FolderPath}\\", StringComparison.OrdinalIgnoreCase))?.StoreId;

            var deletedFolder = folders.FirstOrDefault(folder =>
                folder.FolderType == OutlookFolderType.Deleted
                && (
                    string.IsNullOrWhiteSpace(storeId)
                    || string.Equals(folder.StoreId, storeId, StringComparison.OrdinalIgnoreCase)
                )
                && (
                    string.Equals(addinPath, folder.FolderPath, StringComparison.OrdinalIgnoreCase)
                    || addinPath.StartsWith($"{folder.FolderPath}\\", StringComparison.OrdinalIgnoreCase)
                ));

            return deletedFolder is not null;
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 查詢 folders request 的狀態與分段結果。
        /// </summary>
        [HttpPost("fetch-result-folders")]
        public IActionResult FetchResultFolders([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "fetch_folder_roots");
        }

        [HttpPost("fetch-result-folder-children")]
        public IActionResult FetchResultFolderChildren([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "fetch_folder_children");
        }

        [HttpPost("fetch-result-find-folder")]
        public IActionResult FetchResultFindFolder([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "find_folder");
        }

        [HttpPost("fetch-result-mails")]
        public IActionResult FetchResultMails([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "fetch_mails");
        }

        [HttpPost("fetch-result-mail-body")]
        public IActionResult FetchResultMailBody([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "fetch_mail_body");
        }

        [HttpPost("fetch-result-mail-attachments")]
        public IActionResult FetchResultMailAttachments([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "fetch_mail_attachments");
        }

        [HttpPost("fetch-result-mail-conversation")]
        public IActionResult FetchResultMailConversation([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "fetch_mail_conversation");
        }

        [HttpPost("fetch-result-export-mail-attachment")]
        public IActionResult FetchResultExportMailAttachment([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "export_mail_attachment");
        }

        [HttpPost("fetch-result-rules")]
        public IActionResult FetchResultRules([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "fetch_rules");
        }

        [HttpPost("fetch-result-manage-rule")]
        public IActionResult FetchResultManageRule([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "manage_rule");
        }

        [HttpPost("fetch-result-categories")]
        public IActionResult FetchResultCategories([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "fetch_categories");
        }

        [HttpPost("fetch-result-signalr-ping")]
        public IActionResult FetchResultSignalRPing([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "ping");
        }

        [HttpPost("fetch-result-calendar")]
        public IActionResult FetchResultCalendar([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "fetch_calendar");
        }

        [HttpPost("fetch-result-update-mail-properties")]
        public IActionResult FetchResultUpdateMailProperties([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "update_mail_properties");
        }

        [HttpPost("fetch-result-upsert-category")]
        public IActionResult FetchResultUpsertCategory([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "upsert_category");
        }

        [HttpPost("fetch-result-create-folder")]
        public IActionResult FetchResultCreateFolder([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "create_folder");
        }

        [HttpPost("fetch-result-delete-folder")]
        public IActionResult FetchResultDeleteFolder([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "delete_folder");
        }

        [HttpPost("fetch-result-move-mail")]
        public IActionResult FetchResultMoveMail([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "move_mail");
        }

        [HttpPost("fetch-result-move-mails")]
        public IActionResult FetchResultMoveMails([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "move_mails");
        }

        [HttpPost("fetch-result-delete-mail")]
        public IActionResult FetchResultDeleteMail([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "delete_mail");
        }

        private IActionResult FetchResult(FetchResultRequest req, params string[] expectedTypes)
        {
            req ??= new FetchResultRequest();
            if (string.IsNullOrWhiteSpace(req.RequestId))
                return BadRequest(new { state = "failed", message = "requestId is required." });

            var status = _commandResults.Get(req.RequestId);
            if (status is null)
                return NotFound(new { requestId = req.RequestId, request = "", state = "failed", message = "request not found", next = new FetchResultNext(), data = new { } });

            if (expectedTypes.Length > 0 && !expectedTypes.Contains(status.Type, StringComparer.OrdinalIgnoreCase))
                return BadRequest(new { requestId = req.RequestId, request = RequestName(status.Type), state = "failed", message = "requestId does not match this fetch-result endpoint.", next = new FetchResultNext(), data = new { } });

            var take = Math.Clamp(req.Take <= 0 ? 100 : req.Take, 1, 500);
            var offset = int.TryParse(req.Cursor, out var parsed) && parsed > 0 ? parsed : 0;
            var command = _commandResults.GetRequestCommand(req.RequestId);
            var (data, next) = GetFetchResultData(status.Type, command, offset, take);

            return Ok(new FetchResultResponse
            {
                RequestId = status.CommandId,
                Request = RequestName(status.Type),
                State = ResultState(status.Status),
                Message = status.Message,
                Next = next,
                Data = data,
            });
        }

        private (object Data, FetchResultNext Next) GetFetchResultData(string requestType, PendingCommand? command, int offset, int take)
        {
            switch (requestType)
            {
                case "fetch_folder_roots":
                case "fetch_folder_children":
                case "create_folder":
                case "delete_folder":
                {
                    var snapshot = OutlookFolderPathMapper.ToApiSnapshot(_mailStore.GetFolderSnapshot());
                    var page = Page(snapshot.Folders, offset, take);
                    return (new { stores = snapshot.Stores, folders = page.Items }, page.Next);
                }
                case "find_folder":
                {
                    var snapshot = OutlookFolderPathMapper.ToApiSnapshot(_mailStore.GetFolderSnapshot());
                    var matches = FindFolderMatches(snapshot.Stores, snapshot.Folders, command?.FindFolderRequest);
                    var requestedMaxResults = command?.FindFolderRequest?.MaxResults ?? 20;
                    var maxResults = Math.Clamp(requestedMaxResults <= 0 ? 20 : requestedMaxResults, 1, 100);
                    var page = Page(matches.Take(maxResults).ToList(), offset, take);
                    var pendingDiscoveryTargets = _mailStore.GetPendingFolderDiscoveryTargets().Count;
                    return (new
                    {
                        query = new
                        {
                            name = command?.FindFolderRequest?.Name ?? string.Empty,
                            folderPath = OutlookFolderPathMapper.ToApiPath(command?.FindFolderRequest?.FolderPath ?? string.Empty),
                            folderType = command?.FindFolderRequest?.FolderType ?? string.Empty,
                            storeId = command?.FindFolderRequest?.StoreId ?? string.Empty,
                        },
                        matchCount = matches.Count,
                        isAmbiguous = matches.Count > 1,
                        discoveryComplete = pendingDiscoveryTargets == 0,
                        pendingDiscoveryTargets,
                        folders = page.Items,
                    }, page.Next);
                }
                case "fetch_mails":
                case "update_mail_properties":
                case "move_mail":
                case "move_mails":
                case "delete_mail":
                {
                    var page = Page(OutlookFolderPathMapper.ToApiMails(_mailStore.GetMails()), offset, take);
                    return (new { mails = page.Items }, page.Next);
                }
                case "fetch_mail_body":
                {
                    var mails = OutlookFolderPathMapper.ToApiMails(_mailStore.GetMails());
                    var mail = FindRequestedMail(
                        mails,
                        command?.MailBodyRequest?.MailId,
                        command?.MailBodyRequest?.FolderPath);
                    var result = mail is null ? new List<MailItemDto>() : new List<MailItemDto> { mail };
                    return (new { mails = result }, new FetchResultNext());
                }
                case "fetch_mail_attachments":
                {
                    var mailId = command?.MailAttachmentsRequest?.MailId ?? string.Empty;
                    var attachments = string.IsNullOrWhiteSpace(mailId)
                        ? null
                        : _mailStore.GetMailAttachments(mailId);
                    var apiAttachments = attachments is null
                        ? null
                        : OutlookFolderPathMapper.ToApiAttachments(attachments);
                    var page = Page(apiAttachments?.Attachments ?? new List<MailAttachmentDto>(), offset, take);
                    return (new { mailId, folderPath = apiAttachments?.FolderPath ?? string.Empty, attachments = page.Items }, page.Next);
                }
                case "fetch_mail_conversation":
                {
                    var mailId = command?.MailConversationRequest?.MailId ?? string.Empty;
                    var conversation = string.IsNullOrWhiteSpace(mailId)
                        ? null
                        : _mailStore.GetMailConversation(mailId);
                    var apiConversation = conversation is null
                        ? new MailConversationDto { MailId = mailId, FolderPath = command?.MailConversationRequest?.FolderPath ?? string.Empty }
                        : OutlookFolderPathMapper.ToApiConversation(conversation);
                    var page = Page(apiConversation.Mails, offset, take);
                    return (new
                    {
                        mailId = apiConversation.MailId,
                        folderPath = apiConversation.FolderPath,
                        conversationId = apiConversation.ConversationId,
                        conversationTopic = apiConversation.ConversationTopic,
                        mails = page.Items,
                    }, page.Next);
                }
                case "export_mail_attachment":
                    return (new { }, new FetchResultNext());
                case "fetch_rules":
                case "manage_rule":
                {
                    var page = Page(_mailStore.GetRules(), offset, take);
                    return (new { rules = page.Items }, page.Next);
                }
                case "fetch_categories":
                case "upsert_category":
                {
                    var page = Page(_mailStore.GetCategories(), offset, take);
                    return (new { categories = page.Items }, page.Next);
                }
                case "fetch_calendar":
                {
                    var page = Page(_mailStore.GetCalendarEvents(), offset, take);
                    return (new { calendarEvents = page.Items }, page.Next);
                }
                default:
                    return (new { }, new FetchResultNext());
            }
        }

        private static MailItemDto? FindRequestedMail(List<MailItemDto> mails, string? mailId, string? folderPath)
        {
            if (string.IsNullOrWhiteSpace(mailId))
                return null;

            var apiFolderPath = OutlookFolderPathMapper.ToApiPath(folderPath ?? string.Empty);
            return mails.FirstOrDefault(mail =>
                string.Equals(mail.Id, mailId, StringComparison.OrdinalIgnoreCase)
                && (
                    string.IsNullOrWhiteSpace(apiFolderPath)
                    || string.Equals(mail.FolderPath, apiFolderPath, StringComparison.OrdinalIgnoreCase)
                ));
        }

        private static (List<T> Items, FetchResultNext Next) Page<T>(List<T> source, int offset, int take)
        {
            var total = source.Count;
            var safeOffset = Math.Clamp(offset, 0, total);
            var items = source.Skip(safeOffset).Take(take).ToList();
            var next = safeOffset + items.Count;
            var hasMore = next < total;
            return (items, new FetchResultNext { Cursor = hasMore ? next.ToString() : string.Empty, HasMore = hasMore });
        }

        private static string ResultState(string status)
        {
            return status switch
            {
                "pending" => "running",
                "completed" or "mocked" => "completed",
                "addin_unavailable" => "unavailable",
                "timeout" => "timeout",
                _ => "failed",
            };
        }

        private static string RequestName(string commandType)
        {
            return commandType switch
            {
                "fetch_folder_roots" => "request-folders",
                "fetch_folder_children" => "request-folder-children",
                "find_folder" => "request-find-folder",
                "fetch_mails" => "request-mails",
                "fetch_mail_body" => "request-mail-body",
                "fetch_mail_attachments" => "request-mail-attachments",
                "fetch_mail_conversation" => "request-mail-conversation",
                "export_mail_attachment" => "request-export-mail-attachment",
                "fetch_rules" => "request-rules",
                "manage_rule" => "request-manage-rule",
                "fetch_categories" => "request-categories",
                "ping" => "request-signalr-ping",
                "fetch_calendar" => "request-calendar",
                "update_mail_properties" => "request-update-mail-properties",
                "upsert_category" => "request-upsert-category",
                "create_folder" => "request-create-folder",
                "delete_folder" => "request-delete-folder",
                "move_mail" => "request-move-mail",
                "move_mails" => "request-move-mails",
                "delete_mail" => "request-delete-mail",
                _ => commandType,
            };
        }

        private static bool RequiresFolderCache(PendingCommand cmd)
        {
            return cmd.Type is
                "fetch_mails"
                or "fetch_mail_body"
                or "fetch_mail_attachments"
                or "fetch_mail_conversation"
                or "export_mail_attachment"
                or "update_mail_properties"
                or "manage_rule"
                or "create_folder"
                or "delete_folder"
                or "move_mail"
                or "move_mails"
                or "delete_mail";
        }

        private IActionResult? ValidateRuleRequest(OutlookRuleCommandRequest? req)
        {
            if (req is null)
                return BadRequest(new { status = "missing_rule_request", message = "rule request is required." });

            req.Operation = string.IsNullOrWhiteSpace(req.Operation) ? "upsert" : req.Operation.Trim().ToLowerInvariant();
            req.RuleType = string.IsNullOrWhiteSpace(req.RuleType) ? "receive" : req.RuleType.Trim().ToLowerInvariant();
            req.RuleName = req.RuleName.Trim();
            req.OriginalRuleName = req.OriginalRuleName.Trim();
            req.Conditions ??= new OutlookRuleConditionsRequest();
            req.Actions ??= new OutlookRuleActionsRequest();
            NormalizeRuleList(req.Conditions.SubjectContains);
            NormalizeRuleList(req.Conditions.BodyContains);
            NormalizeRuleList(req.Conditions.SenderAddressContains);
            NormalizeRuleList(req.Conditions.Categories);
            NormalizeRuleList(req.Actions.AssignCategories);
            req.Actions.MoveToFolderPath = OutlookFolderPathMapper.ToAddinPath(req.Actions.MoveToFolderPath.Trim());

            if (req.Operation is not "upsert" and not "delete" and not "set_enabled")
                return BadRequest(new { status = "invalid_rule_operation", message = "operation must be upsert, delete, or set_enabled." });
            if (req.RuleType is not "receive" and not "send")
                return BadRequest(new { status = "invalid_rule_type", message = "ruleType must be receive or send." });
            if (string.IsNullOrWhiteSpace(req.RuleName) && string.IsNullOrWhiteSpace(req.OriginalRuleName))
                return BadRequest(new { status = "missing_rule_name", message = "ruleName or originalRuleName is required." });
            if (req.Conditions.HasAttachment == false)
                return BadRequest(new { status = "unsupported_rule_condition", message = "Outlook object model only supports the has-attachment rule condition." });

            if (req.Operation is "delete" or "set_enabled") return null;

            var hasCondition = req.Conditions.SubjectContains.Count > 0
                || req.Conditions.BodyContains.Count > 0
                || req.Conditions.SenderAddressContains.Count > 0
                || req.Conditions.Categories.Count > 0
                || req.Conditions.HasAttachment is not null;
            var hasAction = !string.IsNullOrWhiteSpace(req.Actions.MoveToFolderPath)
                || req.Actions.AssignCategories.Count > 0
                || req.Actions.MarkAsTask
                || req.Actions.StopProcessingMoreRules;
            if (!hasCondition)
                return BadRequest(new { status = "missing_rule_condition", message = "至少需要一個可由 Outlook object model 建立的條件。" });
            if (!hasAction)
                return BadRequest(new { status = "missing_rule_action", message = "至少需要一個可由 Outlook object model 建立的動作。" });
            return null;
        }

        private static void NormalizeRuleList(List<string> values)
        {
            var normalized = values
                .Select(value => value.Trim())
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            values.Clear();
            values.AddRange(normalized);
        }

        private static List<FolderDto> FindFolderMatches(List<OutlookStoreDto> stores, List<FolderDto> folders, FindFolderRequest? request)
        {
            request ??= new FindFolderRequest();
            var folderPath = OutlookFolderPathMapper.ToApiPath(request.FolderPath);
            var byPath = !string.IsNullOrWhiteSpace(folderPath);
            var name = request.Name.Trim();
            var folderType = OutlookFolderType.Unknown;
            var byFolderType = !string.IsNullOrWhiteSpace(request.FolderType)
                && Enum.TryParse<OutlookFolderType>(request.FolderType, ignoreCase: true, out folderType);
            var primaryStoreId = stores.FirstOrDefault()?.StoreId ?? string.Empty;
            var storeId = request.StoreId;
            if (byFolderType && string.IsNullOrWhiteSpace(storeId))
                storeId = primaryStoreId;

            return folders
                .Where(folder => request.IncludeHidden || !folder.IsHidden)
                .Where(folder => string.IsNullOrWhiteSpace(storeId)
                    || string.Equals(folder.StoreId, storeId, StringComparison.OrdinalIgnoreCase))
                .Where(folder => byPath
                    ? string.Equals(folder.FolderPath, folderPath, StringComparison.OrdinalIgnoreCase)
                    : byFolderType
                        ? folder.FolderType == folderType
                        : string.Equals(folder.Name, name, StringComparison.OrdinalIgnoreCase))
                .OrderBy(folder => folder.StoreId)
                .ThenBy(folder => folder.FolderPath)
                .ToList();
        }

        private Func<bool>? DataReadyPredicate(PendingCommand cmd)
        {
            return cmd.Type switch
            {
                "fetch_folder_roots" => () => _mailStore.CountStoreRoots() > 0,
                "fetch_folder_children" when cmd.FolderDiscoveryRequest is not null => () =>
                    _mailStore.IsFolderChildrenLoaded(
                        cmd.FolderDiscoveryRequest.StoreId,
                        cmd.FolderDiscoveryRequest.ParentEntryId,
                        cmd.FolderDiscoveryRequest.ParentFolderPath),
                _ => null,
            };
        }

        /// <summary>
        /// Web UI 取得 mail list data。
        /// </summary>
        [HttpGet("mails")]
        public IActionResult GetMails()
        {
            return Ok(OutlookFolderPathMapper.ToApiMails(_mailStore.GetMails()));
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 取得單封 mail 的 attachment metadata。
        /// </summary>
        [HttpGet("mail-attachments")]
        public IActionResult GetMailAttachments([FromQuery] string mailId)
        {
            if (string.IsNullOrWhiteSpace(mailId))
                return BadRequest(new { status = "missing_mail_id" });

            var attachments = _mailStore.GetMailAttachments(mailId);
            return attachments is null ? NotFound(new { status = "not_found" }) : Ok(OutlookFolderPathMapper.ToApiAttachments(attachments));
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 取得上次載入的單封 mail conversation。
        /// </summary>
        [HttpGet("mail-conversation")]
        public IActionResult GetMailConversation([FromQuery] string mailId)
        {
            if (string.IsNullOrWhiteSpace(mailId))
                return BadRequest(new { status = "missing_mail_id" });

            var conversation = _mailStore.GetMailConversation(mailId);
            return conversation is null ? NotFound(new { status = "not_found" }) : Ok(OutlookFolderPathMapper.ToApiConversation(conversation));
        }

        /// <summary>
        /// Web UI 取得 folder data。
        /// </summary>
        [HttpGet("folders")]
        public IActionResult GetFolders()
        {
            return Ok(OutlookFolderPathMapper.ToApiSnapshot(_mailStore.GetFolderSnapshot()));
        }

        /// <summary>
        /// Web UI 取得 Outlook rules。
        /// </summary>
        [HttpGet("rules")]
        public IActionResult GetRules()
        {
            return Ok(_mailStore.GetRules());
        }

        /// <summary>
        /// Web UI 取得 Outlook master category list。
        /// </summary>
        [HttpGet("categories")]
        public IActionResult GetCategories()
        {
            return Ok(_mailStore.GetCategories());
        }

        /// <summary>
        /// Web UI 取得 Outlook calendar events。
        /// </summary>
        [HttpGet("calendar")]
        public IActionResult GetCalendar()
        {
            return Ok(_mailStore.GetCalendarEvents());
        }

        /// <summary>
        /// Web 或 Outlook 送出 chat message。
        /// </summary>
        [HttpPost("chat")]
        public async Task<IActionResult> PostChat([FromBody] ChatMessageDto msg, CancellationToken ct)
        {
            if (string.IsNullOrWhiteSpace(msg.Source))
                msg.Source = "web";

            msg.Timestamp = DateTime.Now;
            _chatStore.Add(msg);
            await _hub.Clients.All.SendAsync("NewChatMessage", msg, ct);
            await _mockOutlook.TryReplyToChatAsync(msg, ct);
            return Ok(msg);
        }

        [HttpGet("chat")]
        public IActionResult GetChat()
        {
            return Ok(_chatStore.GetAll());
        }

        /// <summary>
        /// AI / MCP client 查詢 command 執行狀態。
        /// </summary>
        [HttpGet("command-results/{commandId}")]
        public IActionResult GetCommandResult(string commandId)
        {
            var result = _commandResults.Get(commandId);
            if (result is null)
                return NotFound();

            return Ok(result);
        }

        /// <summary>
        /// AI / MCP client 查詢最近 command 執行狀態。
        /// </summary>
        [HttpGet("command-results")]
        public IActionResult GetCommandResults()
        {
            return Ok(_commandResults.GetRecent());
        }

        // ===================== Admin endpoints =====================

        [HttpGet("admin/status")]
        public IActionResult GetAddinStatus()
        {
            return Ok(_addinStatus.GetStatus());
        }

        [HttpGet("admin/logs")]
        public IActionResult GetAddinLogs()
        {
            return Ok(_addinStatus.GetLogs());
        }

        [HttpPost("admin/log")]
        public async Task<IActionResult> PostAddinLog([FromBody] AddinLogEntry entry)
        {
            _addinStatus.AddLog(entry.Level, entry.Message);
            await _hub.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs());
            return Ok();
        }
    }
}
