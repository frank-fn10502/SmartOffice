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
        private readonly OutlookSignalRCommandDispatcher _commandDispatcher;
        private readonly MockOutlookService _mockOutlook;
        private readonly IHubContext<NotificationHub> _hub;
        private readonly AddinStatusStore _addinStatus;
        private readonly AttachmentExportService _attachmentExports;

        public OutlookController(MailStore mailStore, ChatStore chatStore,
            CommandResultStore commandResults, OutlookSignalRCommandDispatcher commandDispatcher, MockOutlookService mockOutlook, IHubContext<NotificationHub> hub, AddinStatusStore addinStatus, AttachmentExportService attachmentExports)
        {
            _mailStore = mailStore;
            _chatStore = chatStore;
            _commandResults = commandResults;
            _commandDispatcher = commandDispatcher;
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
            return await DispatchCommandAsync(new PendingCommand { Type = "fetch_folders" }, ct);
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
            _commandResults.RecordDispatched(cmd);

            if (await _mockOutlook.TryDispatchAsync(cmd, ct))
            {
                _commandResults.RecordResult(new OutlookCommandResult
                {
                    CommandId = cmd.Id,
                    Success = true,
                    Message = $"{cmd.Type} completed by mock backend",
                    Timestamp = DateTime.Now,
                });
                return Ok(new { commandId = cmd.Id, status = "mocked" });
            }

            var dispatched = await _commandDispatcher.DispatchAsync(cmd, ct);
            if (!dispatched)
            {
                _commandResults.RecordUnavailable(cmd);
                return Conflict(new { commandId = cmd.Id, status = "addin_unavailable" });
            }

            await _hub.Clients.All.SendAsync("AddinStatus", _addinStatus.GetStatus(), ct);
            return Ok(new { commandId = cmd.Id, status = "dispatched" });
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
