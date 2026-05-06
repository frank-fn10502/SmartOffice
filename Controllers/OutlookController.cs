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
            req.ParentFolderPath = OutlookFolderPathMapper.ToAddinPath(req.ParentFolderPath);
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
            var result = await _folderCache.FetchFolderRootsQueuedAsync(ct);
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
            var cmd = new PendingCommand
            {
                Type = "delete_mail",
                DeleteMailRequest = req
            };
            return await DispatchCommandAsync(cmd, ct);
        }

        private async Task<IActionResult> DispatchCommandAsync(PendingCommand cmd, CancellationToken ct)
        {
            OutlookFolderPathMapper.NormalizeRequest(cmd);

            // Hub 已收下的 AddIn request 要由中心 queue 控制完成，避免外部 client 斷線造成 Outlook automation 半路取消。
            return await _commandQueue.ExecuteExclusiveAsync(async operationCt =>
            {
                if (RequiresFolderCache(cmd) && !await _folderCache.EnsureFolderCacheAsync(cmd, operationCt))
                {
                    var failedBody = new { commandId = cmd.Id, status = "folder_cache_unavailable", message = "Hub could not load Outlook folders before running folder-dependent command." };
                    return Conflict(failedBody);
                }

                var result = await _commandQueue.ExecuteQueuedCommandAsync(cmd, DataReadyPredicate(cmd), ct: operationCt);
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
                or "update_mail_properties"
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

        /// <summary>
        /// Web UI 取得 cached mails。
        /// </summary>
        [HttpGet("mails")]
        public IActionResult GetMails()
        {
            return Ok(OutlookFolderPathMapper.ToApiMails(_mailStore.GetMails()));
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 取得單封 mail 的 cached attachment metadata。
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
        /// Web UI 取得 cached folders。
        /// </summary>
        [HttpGet("folders")]
        public IActionResult GetFolders()
        {
            return Ok(OutlookFolderPathMapper.ToApiSnapshot(_mailStore.GetFolderSnapshot()));
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
