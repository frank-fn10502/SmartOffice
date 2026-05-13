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
    public partial class OutlookController : ControllerBase
    {
        private const int MaxMoveMailsBatchSize = 500;

        private readonly MailStore _mailStore;
        private readonly CommandResultStore _commandResults;
        private readonly OutlookCommandQueue _commandQueue;
        private readonly IHubContext<NotificationHub> _hub;
        private readonly AddinStatusStore _addinStatus;
        private readonly OutlookFolderCacheService _folderCache;
        private readonly OutlookFetchResultService _fetchResults;

        public OutlookController(MailStore mailStore,
            CommandResultStore commandResults, OutlookCommandQueue commandQueue, IHubContext<NotificationHub> hub, AddinStatusStore addinStatus, OutlookFolderCacheService folderCache, OutlookFetchResultService fetchResults)
        {
            _mailStore = mailStore;
            _commandResults = commandResults;
            _commandQueue = commandQueue;
            _hub = hub;
            _addinStatus = addinStatus;
            _folderCache = folderCache;
            _fetchResults = fetchResults;
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
                    request = RequestName("find_folder"),
                    status = "missing_folder_query",
                    state = "failed",
                    message = "name, folderPath, or folderType is required.",
                    data = new { },
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
                Message = success ? "completed" : "folder_unavailable",
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
            req ??= new FetchCalendarRequest();
            req.DaysForward = Math.Clamp(req.DaysForward <= 0 ? 31 : req.DaysForward, 1, 366);
            req.StartDate = UtcDateTime.Normalize(req.StartDate);
            req.EndDate = UtcDateTime.Normalize(req.EndDate);
            if (req.StartDate.HasValue && req.EndDate.HasValue && req.StartDate > req.EndDate)
            {
                return BadRequest(new
                {
                    request = RequestName("fetch_calendar"),
                    status = "invalid_calendar_range",
                    state = "failed",
                    message = "startDate must be earlier than or equal to endDate.",
                    data = new { },
                });
            }

            var cmd = new PendingCommand
            {
                Type = "fetch_calendar",
                CalendarRequest = req
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 要求 Outlook AddIn 背景同步真正通訊錄。
        /// </summary>
        [HttpPost("request-address-book")]
        public async Task<IActionResult> RequestAddressBook([FromBody] AddressBookSyncRequest? req, CancellationToken ct)
        {
            req ??= new AddressBookSyncRequest();
            req.MaxContacts = Math.Clamp(req.MaxContacts <= 0 ? 1000 : req.MaxContacts, 1, 5000);
            req.MaxAddressEntriesPerList = Math.Clamp(req.MaxAddressEntriesPerList <= 0 ? 500 : req.MaxAddressEntriesPerList, 1, 2000);

            var cmd = new PendingCommand
            {
                Type = "fetch_address_book",
                AddressBookRequest = req
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        [HttpPost("request-update-mail-properties")]
        public async Task<IActionResult> RequestUpdateMailProperties([FromBody] MailPropertiesCommandRequest req, CancellationToken ct)
        {
            var error = ApiRequestValidation.RequireFields(("mailId", req?.MailId), ("folderPath", req?.FolderPath));
            if (error is not null) return error;

            if (!CanUpdateMailProperties(req?.MailId, out var unsupported))
                return UnsupportedMailPropertiesUpdate(unsupported);

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
            var error = ApiRequestValidation.RequireFields(("name", req?.Name));
            if (error is not null) return error;

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
            var error = ApiRequestValidation.RequireFields(("parentFolderPath", req?.ParentFolderPath), ("name", req?.Name));
            if (error is not null) return error;

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
            var error = ApiRequestValidation.RequireFields(("folderPath", req?.FolderPath));
            if (error is not null) return error;

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
            var error = ApiRequestValidation.RequireFields(
                ("mailId", req?.MailId),
                ("sourceFolderPath", req?.SourceFolderPath),
                ("destinationFolderPath", req?.DestinationFolderPath));
            if (error is not null) return error;

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
            var error = ApiRequestValidation.RequireFields(("destinationFolderPath", req?.DestinationFolderPath));
            if (error is not null) return error;
            if (req is null)
                return ApiRequestValidation.MissingRequiredFields("destinationFolderPath", "mailIds");

            req.MailIds ??= new List<string>();
            req.SourceFolderPaths ??= new List<string>();

            if (req.MailIds.Count == 0)
                return ApiRequestValidation.MissingRequiredFields("mailIds");
            if (string.IsNullOrWhiteSpace(req.SourceFolderPath) && req.SourceFolderPaths.Count == 0)
                return ApiRequestValidation.MissingRequiredFields("sourceFolderPath");

            if (req.MailIds.Count > MaxMoveMailsBatchSize)
            {
                return BadRequest(new
                {
                    request = RequestName("move_mails"),
                    status = "too_many_mail_ids",
                    state = "failed",
                    message = $"move_mails 單次最多只能移動 {MaxMoveMailsBatchSize} 封郵件；請由 caller 分批呼叫。",
                    maxBatchSize = MaxMoveMailsBatchSize,
                    actualCount = req.MailIds.Count,
                    data = new { },
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
            var error = ApiRequestValidation.RequireFields(("mailId", req?.MailId), ("folderPath", req?.FolderPath));
            if (error is not null) return error;

            var cmd = new PendingCommand
            {
                Type = "delete_mail",
                DeleteMailRequest = req
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        private bool CanUpdateMailProperties(string? mailId, out MailItemDto? unsupported)
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

        private IActionResult UnsupportedMailPropertiesUpdate(MailItemDto? mail)
        {
            return BadRequest(new
            {
                status = "unsupported_outlook_item_type",
                state = "failed",
                request = RequestName("update_mail_properties"),
                operation = "update_mail_properties",
                mailId = mail?.Id ?? "",
                messageClass = mail?.MessageClass ?? "",
                message = "此 Outlook item 不是一般郵件，不能更新一般郵件屬性；會議邀請可讀取與搬移，但分類、旗標、已讀狀態需要 Outlook 會議/行事曆流程。",
                data = new { },
            });
        }

        private Task<IActionResult> DispatchCommandAsync(PendingCommand cmd, CancellationToken ct)
        {
            OutlookFolderPathMapper.NormalizeRequest(cmd);

            // Hub 已收下的 Outlook request 要由中心 queue 控制完成，避免外部 client 斷線造成 Outlook automation 半路取消。
            _commandResults.RecordDispatched(cmd);
            _ = Task.Run(() => _commandQueue.ExecuteExclusiveAsync(async operationCt =>
            {
                if (RequiresFolderCache(cmd) && !await _folderCache.EnsureFolderCacheAsync(cmd, operationCt))
                {
                    _commandResults.RecordResult(new OutlookCommandResult
                    {
                        CommandId = cmd.Id,
                        Success = false,
                        Message = "folder_unavailable",
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

        private (FetchMailsRequest Request, IActionResult? Error) BuildFetchMailsRequest(RequestMailsApiRequest req)
        {
            req ??= new RequestMailsApiRequest();
            if (string.IsNullOrWhiteSpace(req.FolderPath))
            {
                return (new FetchMailsRequest(), BadRequest(new
                {
                    request = RequestName("fetch_mails"),
                    status = "missing_required_fields",
                    state = "failed",
                    message = "Missing required request field(s): folderPath.",
                    requiredFields = new[] { "folderPath" },
                    data = new { },
                }));
            }

            var request = new FetchMailsRequest
            {
                FolderPath = req.FolderPath,
                ReceivedFrom = UtcDateTime.Normalize(req.ReceivedFrom),
                ReceivedTo = UtcDateTime.Normalize(req.ReceivedTo),
                MaxCount = Math.Clamp(req.MaxCount <= 0 ? 30 : req.MaxCount, 1, 500),
            };
            if (req.LookbackHours is null) return (request, null);

            if (req.LookbackHours <= 0)
            {
                return (request, BadRequest(new
                {
                    request = RequestName("fetch_mails"),
                    status = "invalid_lookback_hours",
                    state = "failed",
                    message = "lookbackHours 必須大於 0，例如 12 代表過去 12 小時、24 代表過去 1 天。",
                    data = new { },
                }));
            }

            if (req.LookbackHours > 24 * 365)
            {
                return (request, BadRequest(new
                {
                    request = RequestName("fetch_mails"),
                    status = "invalid_lookback_hours",
                    state = "failed",
                    message = "lookbackHours 不可超過 8760 小時。若需要更大的範圍，請改用 receivedFrom / receivedTo。",
                    data = new { },
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

        [HttpPost("fetch-result-address-book")]
        public IActionResult FetchResultAddressBook([FromBody] FetchResultRequest req)
        {
            return FetchResult(req, "fetch_address_book");
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
            return _fetchResults.FetchResult(req, expectedTypes);
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

    }
}
