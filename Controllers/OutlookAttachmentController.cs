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
    public class OutlookAttachmentController : ControllerBase
    {
        private readonly MailStore _mailStore;
        private readonly CommandResultStore _commandResults;
        private readonly OutlookCommandQueue _commandQueue;
        private readonly IHubContext<NotificationHub> _hub;
        private readonly AddinStatusStore _addinStatus;
        private readonly AttachmentExportService _attachmentExports;
        private readonly OutlookFolderCacheService _folderCache;

        public OutlookAttachmentController(
            MailStore mailStore,
            CommandResultStore commandResults,
            OutlookCommandQueue commandQueue,
            IHubContext<NotificationHub> hub,
            AddinStatusStore addinStatus,
            AttachmentExportService attachmentExports,
            OutlookFolderCacheService folderCache)
        {
            _mailStore = mailStore;
            _commandResults = commandResults;
            _commandQueue = commandQueue;
            _hub = hub;
            _addinStatus = addinStatus;
            _attachmentExports = attachmentExports;
            _folderCache = folderCache;
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 要求單封 mail body；mail list 本身只應先載入 metadata。
        /// </summary>
        [HttpPost("request-mail-body")]
        public async Task<IActionResult> RequestMailBody([FromBody] FetchMailBodyRequest req, CancellationToken ct)
        {
            var error = ApiRequestValidation.RequireFields(("mailId", req?.MailId), ("folderPath", req?.FolderPath));
            if (error is not null) return error;

            return await DispatchCommandAsync(new PendingCommand
            {
                Type = "fetch_mail_body",
                MailBodyRequest = req
            }, ct);
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 要求單封 mail attachment metadata。
        /// </summary>
        [HttpPost("request-mail-attachments")]
        public async Task<IActionResult> RequestMailAttachments([FromBody] FetchMailAttachmentsRequest req, CancellationToken ct)
        {
            var error = ApiRequestValidation.RequireFields(("mailId", req?.MailId), ("folderPath", req?.FolderPath));
            if (error is not null) return error;

            return await DispatchCommandAsync(new PendingCommand
            {
                Type = "fetch_mail_attachments",
                MailAttachmentsRequest = req
            }, ct);
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 要求指定 mail 所屬 Outlook conversation。
        /// </summary>
        [HttpPost("request-mail-conversation")]
        public async Task<IActionResult> RequestMailConversation([FromBody] FetchMailConversationRequest req, CancellationToken ct)
        {
            req ??= new FetchMailConversationRequest();
            var error = ApiRequestValidation.RequireFields(("mailId", req.MailId), ("folderPath", req.FolderPath));
            if (error is not null) return error;

            req.MaxCount = Math.Clamp(req.MaxCount <= 0 ? 100 : req.MaxCount, 1, 300);
            return await DispatchCommandAsync(new PendingCommand
            {
                Type = "fetch_mail_conversation",
                MailConversationRequest = req
            }, ct);
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 要求 AddIn 將指定 attachment 匯出到本機約定目錄。
        /// </summary>
        [HttpPost("request-export-mail-attachment")]
        public async Task<IActionResult> RequestExportMailAttachment([FromBody] ExportMailAttachmentRequest req, CancellationToken ct)
        {
            var error = ApiRequestValidation.RequireFields(
                ("mailId", req?.MailId),
                ("folderPath", req?.FolderPath),
                ("attachmentId", req?.AttachmentId));
            if (error is not null) return error;

            if (string.IsNullOrWhiteSpace(req!.ExportRootPath))
                req.ExportRootPath = _attachmentExports.RootPath;

            return await DispatchCommandAsync(new PendingCommand
            {
                Type = "export_mail_attachment",
                ExportMailAttachmentRequest = req
            }, ct);
        }

        /// <summary>
        /// Web UI Host 開啟已匯出的附件；只接受 Hub 已記錄的 exported attachment id。
        /// </summary>
        [HttpPost("open-exported-attachment")]
        public IActionResult OpenExportedAttachment([FromBody] OpenExportedAttachmentRequest req)
        {
            var error = ApiRequestValidation.RequireFields(("exportedAttachmentId", req?.ExportedAttachmentId));
            if (error is not null) return error;

            if (!_mailStore.TryGetExportedAttachment(req!.ExportedAttachmentId, out var attachment))
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
            var error = ApiRequestValidation.RequireFields(("rootPath", req?.RootPath));
            if (error is not null) return error;

            try
            {
                return Ok(_attachmentExports.UpdateSettings(req!.RootPath));
            }
            catch (Exception ex)
            {
                return BadRequest(new { status = "update_failed", message = ex.Message });
            }
        }

        private Task<IActionResult> DispatchCommandAsync(PendingCommand cmd, CancellationToken ct)
        {
            OutlookFolderPathMapper.NormalizeRequest(cmd);
            _commandResults.RecordDispatched(cmd);
            _ = Task.Run(() => _commandQueue.ExecuteExclusiveAsync(async operationCt =>
            {
                if (!await _folderCache.EnsureFolderCacheAsync(cmd, operationCt))
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

                await _commandQueue.ExecuteQueuedCommandAsync(cmd, ct: operationCt);
                await _hub.Clients.All.SendAsync("AddinStatus", _addinStatus.GetStatus(), CancellationToken.None);
                return true;
            }, CancellationToken.None));

            return Task.FromResult<IActionResult>(Ok(OperationAccepted(cmd)));
        }

        private static object OperationAccepted(PendingCommand command)
        {
            return new
            {
                requestId = command.Id,
                request = RequestName(command.Type),
                state = "accepted",
                message = "Request accepted. Poll the paired fetch-result-* endpoint for state and data.",
                data = new { },
            };
        }

        private static string RequestName(string commandType)
        {
            return commandType switch
            {
                "fetch_mail_body" => "request-mail-body",
                "fetch_mail_attachments" => "request-mail-attachments",
                "fetch_mail_conversation" => "request-mail-conversation",
                "export_mail_attachment" => "request-export-mail-attachment",
                _ => commandType,
            };
        }
    }
}
