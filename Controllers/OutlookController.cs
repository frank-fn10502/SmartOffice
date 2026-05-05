using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.SignalR;
using SmartOffice.Hub.Hubs;
using SmartOffice.Hub.Models;
using SmartOffice.Hub.Services;

namespace SmartOffice.Hub.Controllers
{
    [ApiController]
    [Route("api/outlook")]
    public class OutlookController : ControllerBase
    {
        private readonly MailStore _mailStore;
        private readonly ChatStore _chatStore;
        private readonly CommandResultStore _commandResults;
        private readonly OutlookCommandQueue _commandQueue;
        private readonly MockOutlookService _mockOutlook;
        private readonly IHubContext<NotificationHub> _hub;
        private readonly AddinStatusStore _addinStatus;
        private readonly AttachmentExportService _attachmentExports;
        private static readonly TimeSpan FolderCacheMaxAge = TimeSpan.FromMinutes(30);

        public OutlookController(MailStore mailStore, ChatStore chatStore,
            CommandResultStore commandResults, OutlookCommandQueue commandQueue, MockOutlookService mockOutlook, IHubContext<NotificationHub> hub, AddinStatusStore addinStatus, AttachmentExportService attachmentExports)
        {
            _mailStore = mailStore;
            _chatStore = chatStore;
            _commandResults = commandResults;
            _commandQueue = commandQueue;
            _mockOutlook = mockOutlook;
            _hub = hub;
            _addinStatus = addinStatus;
            _attachmentExports = attachmentExports;
        }

        // ===================== Web UI 呼叫這些 endpoint =====================

        /// <summary>
        /// Web UI、AI 或 MCP client 要求 mails；Hub 會透過 SignalR dispatch 給 Outlook AddIn。
        /// </summary>
        [HttpPost("request-mails")]
        public async Task<IActionResult> RequestMails([FromBody] FetchMailsRequest req, CancellationToken ct)
        {
            var cmd = new PendingCommand
            {
                Type = "fetch_mails",
                MailsRequest = req
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 要求搜尋 mails；Hub 會展開 folder scope 並分片 dispatch 給 Outlook AddIn。
        /// </summary>
        [HttpPost("request-mail-search")]
        public async Task<IActionResult> RequestMailSearch([FromBody] SearchMailsRequest req, CancellationToken ct)
        {
            req ??= new SearchMailsRequest();
            if (string.IsNullOrWhiteSpace(req.SearchId))
                req.SearchId = Guid.NewGuid().ToString();
            NormalizeMailSearchRequest(req);

            var cmd = new PendingCommand
            {
                Type = "search_mails",
                SearchMailsRequest = req
            };
            return await _commandQueue.ExecuteExclusiveAsync(
                operationCt => RequestMailSearchQueuedAsync(cmd, req, operationCt),
                CancellationToken.None);
        }

        private async Task<IActionResult> RequestMailSearchQueuedAsync(PendingCommand cmd, SearchMailsRequest req, CancellationToken ct)
        {
            var folderReady = await EnsureFolderCacheForMailSearchAsync(cmd, ct);
            if (!folderReady)
            {
                _commandResults.RecordDispatched(cmd);
                _commandResults.RecordResult(new OutlookCommandResult
                {
                    CommandId = cmd.Id,
                    Success = false,
                    Message = "Hub could not load Outlook folders for mail search.",
                    Timestamp = DateTime.Now,
                });
                var failed = _mailStore.UpdateMailSearchProgress(new MailSearchProgressDto
                {
                    SearchId = req.SearchId,
                    CommandId = cmd.Id,
                    Status = "failed",
                    Phase = "load_folders",
                    Message = "folder_cache_unavailable",
                    Timestamp = DateTime.Now,
                });
                await _hub.Clients.All.SendAsync("MailSearchProgress", failed, ct);
                return Conflict(new { commandId = cmd.Id, searchId = req.SearchId, status = "folder_cache_unavailable" });
            }

            var slices = BuildMailSearchSlices(req, cmd.Id);
            if (slices.Count == 0)
            {
                _commandResults.RecordDispatched(cmd);
                _commandResults.RecordResult(new OutlookCommandResult
                {
                    CommandId = cmd.Id,
                    Success = false,
                    Message = "Hub loaded folders but no searchable folder matched the request.",
                    Timestamp = DateTime.Now,
                });
                var failed = _mailStore.UpdateMailSearchProgress(new MailSearchProgressDto
                {
                    SearchId = req.SearchId,
                    CommandId = cmd.Id,
                    Status = "failed",
                    Phase = "planning",
                    Message = "no_searchable_folder",
                    Timestamp = DateTime.Now,
                });
                await _hub.Clients.All.SendAsync("MailSearchProgress", failed, ct);
                return Conflict(new { commandId = cmd.Id, searchId = req.SearchId, status = "no_searchable_folder" });
            }

            req.ScopeFolderPaths = slices.Select(slice => slice.FolderPath).ToList();
            var progress = _mailStore.StartMailSearchProgress(req, cmd.Id);
            await _hub.Clients.All.SendAsync("MailSearchProgress", progress, ct);
            _commandResults.RecordDispatched(cmd);
            var result = await DispatchMailSearchSlicesAsync(cmd, slices, ct);
            var body = new
            {
                commandId = cmd.Id,
                searchId = req.SearchId,
                status = result.Status,
                message = result.Message,
                sliceCount = slices.Count,
            };
            return result.Success ? Ok(body) : StatusCode(result.HttpStatusCode, body);
        }

        private void NormalizeMailSearchRequest(SearchMailsRequest req)
        {
            req.ScopeFolderPaths ??= new List<string>();
            req.TextFields = NormalizeMailSearchTextFields(req.TextFields);
            req.CategoryNames = NormalizeMailSearchCategoryNames(req.CategoryNames);
            req.FlagState = NormalizeMailSearchState(req.FlagState, new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "any", "flagged", "unflagged" });
            req.ReadState = NormalizeMailSearchState(req.ReadState, new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "any", "unread", "read" });
        }

        private List<MailSearchSliceRequest> BuildMailSearchSlices(SearchMailsRequest req, string parentCommandId)
        {
            var snapshot = _mailStore.GetFolderSnapshot();
            var requestedPaths = req.ScopeFolderPaths
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            var searchableFolders = snapshot.Folders
                .Where(folder => !folder.IsStoreRoot)
                .Where(folder => string.IsNullOrWhiteSpace(req.StoreId) || string.Equals(folder.StoreId, req.StoreId, StringComparison.OrdinalIgnoreCase))
                .ToList();

            var scopedFolders = requestedPaths.Count == 0
                ? searchableFolders
                : searchableFolders
                    .Where(folder => requestedPaths.Any(path =>
                        string.Equals(folder.FolderPath, path, StringComparison.OrdinalIgnoreCase)
                        || (req.IncludeSubFolders && folder.FolderPath.StartsWith($"{path}\\", StringComparison.OrdinalIgnoreCase))))
                    .ToList();

            var plannedFolders = scopedFolders
                .Where(folder => !string.IsNullOrWhiteSpace(folder.FolderPath))
                .DistinctBy(folder => folder.FolderPath, StringComparer.OrdinalIgnoreCase)
                .OrderBy(folder => folder.StoreId)
                .ThenBy(folder => folder.FolderPath)
                .ToList();

            return plannedFolders
                .Select((folder, index) => new MailSearchSliceRequest
                {
                    SearchId = req.SearchId,
                    ParentCommandId = parentCommandId,
                    StoreId = folder.StoreId,
                    FolderPath = folder.FolderPath,
                    Keyword = req.Keyword,
                    TextFields = new List<string>(req.TextFields),
                    CategoryNames = new List<string>(req.CategoryNames),
                    HasAttachments = req.HasAttachments,
                    FlagState = req.FlagState,
                    ReadState = req.ReadState,
                    ReceivedFrom = req.ReceivedFrom,
                    ReceivedTo = req.ReceivedTo,
                    SliceIndex = index,
                    SliceCount = plannedFolders.Count,
                    ResetSearchResults = index == 0,
                    CompleteSearchOnSlice = index == plannedFolders.Count - 1,
                })
                .ToList();
        }

        private static List<string> NormalizeMailSearchTextFields(List<string>? textFields)
        {
            var allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "subject",
                "sender",
                "body"
            };
            var normalized = (textFields ?? new List<string>())
                .Where(field => allowed.Contains(field))
                .Select(field => field.ToLowerInvariant())
                .Distinct()
                .ToList();
            return normalized.Count > 0 ? normalized : new List<string> { "subject" };
        }

        private static List<string> NormalizeMailSearchCategoryNames(List<string>? categoryNames)
        {
            return (categoryNames ?? new List<string>())
                .Select(category => category.Trim())
                .Where(category => !string.IsNullOrWhiteSpace(category))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static string NormalizeMailSearchState(string state, HashSet<string> allowed)
        {
            var normalized = (state ?? string.Empty).Trim().ToLowerInvariant();
            return allowed.Contains(normalized) ? normalized : "any";
        }

        private async Task<bool> EnsureFolderCacheForMailSearchAsync(PendingCommand searchCommand, CancellationToken ct)
        {
            return await EnsureFolderCacheAsync(searchCommand, ct);
        }

        private async Task<bool> EnsureFolderCacheAsync(PendingCommand? ownerCommand, CancellationToken ct)
        {
            if (_mailStore.CountFolders() <= 0 || _mailStore.IsFolderCacheStale(FolderCacheMaxAge))
            {
                await SendFolderCacheProgressAsync(ownerCommand, "Hub is loading Outlook folders before running folder-dependent command.", ct);

                var result = await FetchFolderRootsQueuedAsync(ct);
                if (!result.Success) return false;
            }

            await LoadPendingFolderDiscoveryTargetsAsync(ownerCommand, ct);
            return _mailStore.CountFolders() > 0;
        }

        private async Task LoadPendingFolderDiscoveryTargetsAsync(PendingCommand? ownerCommand, CancellationToken ct)
        {
            const int maxDiscoveryCommands = 200;
            var dispatched = 0;
            while (dispatched < maxDiscoveryCommands)
            {
                var target = _mailStore.GetPendingFolderDiscoveryTargets().FirstOrDefault();
                if (target is null) return;

                await SendFolderCacheProgressAsync(ownerCommand, $"Hub is loading folder children before running folder-dependent command: {target.FolderPath}", ct, target);

                var command = new PendingCommand
                {
                    Type = "fetch_folder_children",
                    FolderDiscoveryRequest = new FolderDiscoveryRequest
                    {
                        StoreId = target.StoreId,
                        ParentEntryId = target.EntryId,
                        ParentFolderPath = target.FolderPath,
                        MaxDepth = 1,
                        MaxChildren = 200,
                        Reset = false,
                    },
                };
                var result = await _commandQueue.ExecuteQueuedCommandAsync(
                    command,
                    () => _mailStore.IsFolderChildrenLoaded(target.StoreId, target.EntryId, target.FolderPath),
                    ensureReady: true,
                    ct: ct);
                if (!result.Success) return;
                dispatched++;
            }
        }

        private async Task SendFolderCacheProgressAsync(PendingCommand? ownerCommand, string message, CancellationToken ct, FolderDto? target = null)
        {
            if (ownerCommand?.Type != "search_mails" || ownerCommand.SearchMailsRequest is null) return;
            var progress = _mailStore.UpdateMailSearchProgress(new MailSearchProgressDto
            {
                SearchId = ownerCommand.SearchMailsRequest.SearchId,
                CommandId = ownerCommand.Id,
                Status = "running",
                Phase = "load_folders",
                CurrentStoreId = target?.StoreId ?? string.Empty,
                CurrentFolderPath = target?.FolderPath ?? string.Empty,
                Message = message,
                Timestamp = DateTime.Now,
            });
            await _hub.Clients.All.SendAsync("MailSearchProgress", progress, ct);
        }

        private async Task<OutlookQueuedCommandResult> DispatchMailSearchSlicesAsync(PendingCommand parentCommand, List<MailSearchSliceRequest> slices, CancellationToken ct)
        {
            const int sliceDelayMs = 100;
            try
            {
                foreach (var slice in slices)
                {
                    var sliceCommand = new PendingCommand
                    {
                        Type = "fetch_mail_search_slice",
                        MailSearchSliceRequest = slice,
                    };
                    slice.CommandId = sliceCommand.Id;
                    var progress = _mailStore.UpdateMailSearchProgress(new MailSearchProgressDto
                    {
                        SearchId = slice.SearchId,
                        CommandId = parentCommand.Id,
                        Status = "running",
                        Phase = "folder",
                        ProcessedFolders = slice.SliceIndex,
                        TotalFolders = slice.SliceCount,
                        CurrentStoreId = slice.StoreId,
                        CurrentFolderPath = slice.FolderPath,
                        ResultCount = _mailStore.GetMailSearchResults().Count,
                        Message = $"Dispatching mail search slice {slice.SliceIndex + 1}/{slice.SliceCount}.",
                        Timestamp = DateTime.Now,
                    });
                    await _hub.Clients.All.SendAsync("MailSearchProgress", progress, ct);

                    var sliceResult = await _commandQueue.ExecuteQueuedCommandAsync(sliceCommand, ensureReady: true, ct: ct);
                    if (!sliceResult.Success)
                    {
                        await CompleteSearchProgressAsync(parentCommand, sliceResult.Status, sliceResult.Message, ct);
                        var result = new OutlookCommandResult
                        {
                            CommandId = parentCommand.Id,
                            Success = false,
                            Message = sliceResult.Message,
                            Timestamp = DateTime.Now,
                        };
                        _commandResults.RecordResult(result);
                        await _hub.Clients.All.SendAsync("CommandResult", result, ct);
                        return OutlookQueuedCommandResult.Failed(parentCommand.Id, sliceResult.Status, sliceResult.Message);
                    }

                    progress = _mailStore.UpdateMailSearchProgress(new MailSearchProgressDto
                    {
                        SearchId = slice.SearchId,
                        CommandId = parentCommand.Id,
                        Status = "running",
                        Phase = "folder",
                        ProcessedFolders = slice.SliceIndex + 1,
                        TotalFolders = slice.SliceCount,
                        CurrentStoreId = slice.StoreId,
                        CurrentFolderPath = slice.FolderPath,
                        ResultCount = _mailStore.GetMailSearchResults().Count,
                        Message = $"Completed mail search slice {slice.SliceIndex + 1}/{slice.SliceCount}.",
                        Timestamp = DateTime.Now,
                    });
                    await _hub.Clients.All.SendAsync("MailSearchProgress", progress, ct);
                    await Task.Delay(sliceDelayMs, ct);
                }

                await CompleteSearchProgressAsync(parentCommand, "completed", "Mail search completed by Hub folder slices.", ct);
                var completed = new OutlookCommandResult
                {
                    CommandId = parentCommand.Id,
                    Success = true,
                    Message = "Mail search completed by Hub folder slices.",
                    Timestamp = DateTime.Now,
                };
                _commandResults.RecordResult(completed);
                await _hub.Clients.All.SendAsync("CommandResult", completed, ct);
                return OutlookQueuedCommandResult.Completed(parentCommand.Id, "completed", completed.Message);
            }
            catch (Exception ex)
            {
                var result = new OutlookCommandResult
                {
                    CommandId = parentCommand.Id,
                    Success = false,
                    Message = $"Mail search dispatch failed: {ex.Message}",
                    Timestamp = DateTime.Now,
                };
                _commandResults.RecordResult(result);
                await CompleteSearchProgressAsync(parentCommand, "failed", result.Message, CancellationToken.None);
                await _hub.Clients.All.SendAsync("CommandResult", result, CancellationToken.None);
                return OutlookQueuedCommandResult.Failed(parentCommand.Id, "failed", result.Message);
            }
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
        public async Task<IActionResult> RequestFolders(CancellationToken ct)
        {
            return await _commandQueue.ExecuteExclusiveAsync(
                operationCt => RequestFoldersQueuedAsync(operationCt),
                CancellationToken.None);
        }

        [HttpPost("request-folder-children")]
        public async Task<IActionResult> RequestFolderChildren([FromBody] FolderDiscoveryRequest req, CancellationToken ct)
        {
            req ??= new FolderDiscoveryRequest();
            req.MaxDepth = Math.Clamp(req.MaxDepth <= 0 ? 1 : req.MaxDepth, 1, 3);
            req.MaxChildren = Math.Clamp(req.MaxChildren <= 0 ? 50 : req.MaxChildren, 1, 200);
            req.Reset = false;

            return await _commandQueue.ExecuteExclusiveAsync(
                operationCt => RequestFolderChildrenQueuedAsync(req, operationCt),
                CancellationToken.None);
        }

        private async Task<IActionResult> RequestFolderChildrenQueuedAsync(FolderDiscoveryRequest req, CancellationToken ct)
        {
            var command = new PendingCommand
            {
                Type = "fetch_folder_children",
                FolderDiscoveryRequest = req,
            };
            var result = await _commandQueue.ExecuteQueuedCommandAsync(
                command,
                () => _mailStore.IsFolderChildrenLoaded(req.StoreId, req.ParentEntryId, req.ParentFolderPath),
                ensureReady: true,
                ct: ct);
            var body = new { commandId = command.Id, status = result.Status, message = result.Message };
            return result.Success ? Ok(body) : StatusCode(result.HttpStatusCode, body);
        }

        private async Task<IActionResult> RequestFoldersQueuedAsync(CancellationToken ct)
        {
            var result = await FetchFolderRootsQueuedAsync(ct);
            var snapshot = _mailStore.GetFolderSnapshot();
            var body = new
            {
                commandId = result.CommandId,
                status = result.Status,
                message = result.Message,
                stores = snapshot.Stores.Count,
                folders = snapshot.Folders.Count,
            };
            return result.Success ? Ok(body) : StatusCode(result.HttpStatusCode, body);
        }

        private async Task<OutlookQueuedCommandResult> FetchFolderRootsQueuedAsync(CancellationToken ct)
        {
            var syncId = Guid.NewGuid().ToString();
            var rootCommand = new PendingCommand
            {
                Type = "fetch_folder_roots",
                FolderDiscoveryRequest = new FolderDiscoveryRequest
                {
                    SyncId = syncId,
                    Reset = true,
                    MaxDepth = 0,
                    MaxChildren = 50,
                },
            };

            var rootResult = await _commandQueue.ExecuteQueuedCommandAsync(
                rootCommand,
                () => _mailStore.CountStoreRoots() > 0,
                ensureReady: true,
                ct: ct);
            if (!rootResult.Success) return rootResult;

            return OutlookQueuedCommandResult.Completed(rootCommand.Id, "completed", "Hub loaded Outlook stores and root folders.");
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 要求 Outlook rule list。
        /// </summary>
        [HttpPost("request-rules")]
        public async Task<IActionResult> RequestRules(CancellationToken ct)
        {
            return await DispatchCommandAsync(new PendingCommand { Type = "fetch_rules" }, ct);
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

        [HttpPost("request-mark-mail-read")]
        public Task<IActionResult> RequestMarkMailRead([FromBody] MailMarkerCommandRequest req, CancellationToken ct)
        {
            return DispatchMailMarkerCommandAsync("mark_mail_read", req, ct);
        }

        [HttpPost("request-mark-mail-unread")]
        public Task<IActionResult> RequestMarkMailUnread([FromBody] MailMarkerCommandRequest req, CancellationToken ct)
        {
            return DispatchMailMarkerCommandAsync("mark_mail_unread", req, ct);
        }

        [HttpPost("request-mark-mail-task")]
        public Task<IActionResult> RequestMarkMailTask([FromBody] MailMarkerCommandRequest req, CancellationToken ct)
        {
            return DispatchMailMarkerCommandAsync("mark_mail_task", req, ct);
        }

        [HttpPost("request-clear-mail-task")]
        public Task<IActionResult> RequestClearMailTask([FromBody] MailMarkerCommandRequest req, CancellationToken ct)
        {
            return DispatchMailMarkerCommandAsync("clear_mail_task", req, ct);
        }

        [HttpPost("request-set-mail-categories")]
        public Task<IActionResult> RequestSetMailCategories([FromBody] MailMarkerCommandRequest req, CancellationToken ct)
        {
            return DispatchMailMarkerCommandAsync("set_mail_categories", req, ct);
        }

        [HttpPost("request-update-mail-properties")]
        public async Task<IActionResult> RequestUpdateMailProperties([FromBody] MailPropertiesCommandRequest req, CancellationToken ct)
        {
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
            var cmd = new PendingCommand
            {
                Type = "move_mail",
                MoveMailRequest = req
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        [HttpPost("request-delete-mail")]
        public async Task<IActionResult> RequestDeleteMail([FromBody] DeleteMailRequest req, CancellationToken ct)
        {
            var cmd = new PendingCommand
            {
                Type = "delete_mail",
                DeleteMailRequest = req
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        private Task<IActionResult> DispatchMailMarkerCommandAsync(string type, MailMarkerCommandRequest req, CancellationToken ct)
        {
            var cmd = new PendingCommand
            {
                Type = type,
                MailMarkerRequest = req
            };
            return DispatchCommandAsync(cmd, ct);
        }

        private async Task<IActionResult> DispatchCommandAsync(PendingCommand cmd, CancellationToken ct)
        {
            // Hub 已收下的 AddIn request 要由中心 queue 控制完成，避免外部 client 斷線造成 Outlook automation 半路取消。
            return await _commandQueue.ExecuteExclusiveAsync(async operationCt =>
            {
                if (RequiresFolderCache(cmd) && !await EnsureFolderCacheAsync(cmd, operationCt))
                {
                    var failedBody = new { commandId = cmd.Id, status = "folder_cache_unavailable", message = "Hub could not load Outlook folders before running folder-dependent command." };
                    return Conflict(failedBody);
                }

                var result = await _commandQueue.ExecuteQueuedCommandAsync(cmd, DataReadyPredicate(cmd), ct: operationCt);
                await CompleteSearchProgressAsync(cmd, result.Success ? "completed" : result.Status, result.Message, CancellationToken.None);
                await _hub.Clients.All.SendAsync("AddinStatus", _addinStatus.GetStatus(), CancellationToken.None);
                var body = new { commandId = result.CommandId == string.Empty ? cmd.Id : result.CommandId, status = result.Status, message = result.Message };
                return result.Success ? Ok(body) : StatusCode(result.HttpStatusCode, body);
            }, CancellationToken.None);
        }

        private static bool RequiresFolderCache(PendingCommand cmd)
        {
            return cmd.Type is
                "fetch_mails"
                or "fetch_mail_body"
                or "fetch_mail_attachments"
                or "export_mail_attachment"
                or "mark_mail_read"
                or "mark_mail_unread"
                or "mark_mail_task"
                or "clear_mail_task"
                or "set_mail_categories"
                or "update_mail_properties"
                or "create_folder"
                or "delete_folder"
                or "move_mail"
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

        private async Task CompleteSearchProgressAsync(PendingCommand cmd, string status, string message, CancellationToken ct)
        {
            if (cmd.Type != "search_mails" || cmd.SearchMailsRequest is null) return;
            var progress = _mailStore.GetMailSearchProgress(cmd.SearchMailsRequest.SearchId) ?? new MailSearchProgressDto
            {
                SearchId = cmd.SearchMailsRequest.SearchId,
                CommandId = cmd.Id,
            };
            progress.Status = status;
            progress.Phase = status == "completed" ? "completed" : "dispatch";
            progress.ResultCount = _mailStore.GetMailSearchResults().Count;
            progress.Message = message;
            progress.Timestamp = DateTime.Now;
            progress = _mailStore.UpdateMailSearchProgress(progress);
            await _hub.Clients.All.SendAsync("MailSearchProgress", progress, ct);
        }

        /// <summary>
        /// Web UI 取得 cached mails。
        /// </summary>
        [HttpGet("mails")]
        public IActionResult GetMails()
        {
            return Ok(_mailStore.GetMails());
        }

        /// <summary>
        /// Web UI 取得 cached mail search results。
        /// </summary>
        [HttpGet("mail-search")]
        public IActionResult GetMailSearchResults()
        {
            return Ok(_mailStore.GetMailSearchResults());
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 查詢 mail search 進度。
        /// </summary>
        [HttpGet("mail-search/progress/{searchId}")]
        public IActionResult GetMailSearchProgress(string searchId)
        {
            var progress = _mailStore.GetMailSearchProgress(searchId);
            return progress is null ? NotFound() : Ok(progress);
        }

        /// <summary>
        /// AI / MCP client 用 command id 查詢對應的 mail search 進度。
        /// </summary>
        [HttpGet("mail-search/progress/by-command/{commandId}")]
        public IActionResult GetMailSearchProgressByCommandId(string commandId)
        {
            var progress = _mailStore.GetMailSearchProgressByCommandId(commandId);
            return progress is null ? NotFound() : Ok(progress);
        }

        /// <summary>
        /// Web UI 取得 cached folders。
        /// </summary>
        [HttpGet("folders")]
        public IActionResult GetFolders()
        {
            return Ok(_mailStore.GetFolderSnapshot());
        }

        /// <summary>
        /// Web UI 取得 cached Outlook rules。
        /// </summary>
        [HttpGet("rules")]
        public IActionResult GetRules()
        {
            return Ok(_mailStore.GetRules());
        }

        /// <summary>
        /// Web UI 取得 cached Outlook master category list。
        /// </summary>
        [HttpGet("categories")]
        public IActionResult GetCategories()
        {
            return Ok(_mailStore.GetCategories());
        }

        /// <summary>
        /// Web UI 取得 cached Outlook calendar events。
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
