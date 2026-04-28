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
        private readonly CommandQueue _commandQueue;
        private readonly IHubContext<NotificationHub> _hub;
        private readonly AddinStatusStore _addinStatus;

        public OutlookController(MailStore mailStore, ChatStore chatStore,
            CommandQueue commandQueue, IHubContext<NotificationHub> hub, AddinStatusStore addinStatus)
        {
            _mailStore = mailStore;
            _chatStore = chatStore;
            _commandQueue = commandQueue;
            _hub = hub;
            _addinStatus = addinStatus;
        }

        // ===================== Web UI calls these =====================

        /// <summary>
        /// Web UI, AI, or MCP client requests mails. Hub queues a command for Outlook Add-in.
        /// </summary>
        [HttpPost("request-mails")]
        public IActionResult RequestMails([FromBody] FetchMailsRequest req)
        {
            var cmd = new PendingCommand
            {
                Type = "fetch_mails",
                MailsRequest = req
            };
            _commandQueue.Enqueue(cmd);
            return Ok(new { commandId = cmd.Id, status = "queued" });
        }

        /// <summary>
        /// Web UI, AI, or MCP client requests folder list.
        /// </summary>
        [HttpPost("request-folders")]
        public IActionResult RequestFolders()
        {
            _commandQueue.Enqueue(new PendingCommand { Type = "fetch_folders" });
            return Ok(new { status = "queued" });
        }

        /// <summary>
        /// Web UI gets cached mails.
        /// </summary>
        [HttpGet("mails")]
        public IActionResult GetMails()
        {
            return Ok(_mailStore.GetMails());
        }

        /// <summary>
        /// Web UI gets cached folders.
        /// </summary>
        [HttpGet("folders")]
        public IActionResult GetFolders()
        {
            return Ok(_mailStore.GetFolders());
        }

        /// <summary>
        /// Web or Outlook sends chat message.
        /// </summary>
        [HttpPost("chat")]
        public async Task<IActionResult> PostChat([FromBody] ChatMessageDto msg)
        {
            msg.Timestamp = DateTime.Now;
            _chatStore.Add(msg);
            await _hub.Clients.All.SendAsync("NewChatMessage", msg);
            return Ok(msg);
        }

        [HttpGet("chat")]
        public IActionResult GetChat()
        {
            return Ok(_chatStore.GetAll());
        }

        // ===================== Outlook Add-in calls these =====================

        /// <summary>
        /// Outlook Add-in long-polls for pending commands (30s timeout).
        /// </summary>
        [HttpGet("poll")]
        public async Task<IActionResult> Poll(CancellationToken ct)
        {
            // Office 2016 add-ins are easiest to deploy behind restricted networks
            // when the desktop process initiates outbound HTTP instead of requiring
            // an inbound socket from the Hub.
            var cmd = await _commandQueue.DequeueAsync(TimeSpan.FromSeconds(30), ct);
            _addinStatus.RecordPoll(cmd?.Type);
            await _hub.Clients.All.SendAsync("AddinStatus", _addinStatus.GetStatus());
            if (cmd == null)
                return Ok(new { type = "none" });
            return Ok(cmd);
        }

        /// <summary>
        /// Outlook Add-in pushes mail results.
        /// </summary>
        [HttpPost("push-mails")]
        public async Task<IActionResult> PushMails([FromBody] List<MailItemDto> mails)
        {
            _mailStore.SetMails(mails);
            _addinStatus.RecordPush("mails", mails.Count);
            await _hub.Clients.All.SendAsync("MailsUpdated", mails);
            await _hub.Clients.All.SendAsync("AddinStatus", _addinStatus.GetStatus());
            await _hub.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs());
            return Ok(new { count = mails.Count });
        }

        /// <summary>
        /// Outlook Add-in pushes folder list.
        /// </summary>
        [HttpPost("push-folders")]
        public async Task<IActionResult> PushFolders([FromBody] List<FolderDto> folders)
        {
            _mailStore.SetFolders(folders);
            _addinStatus.RecordPush("folders", folders.Count);
            await _hub.Clients.All.SendAsync("FoldersUpdated", folders);
            await _hub.Clients.All.SendAsync("AddinStatus", _addinStatus.GetStatus());
            await _hub.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs());
            return Ok(new { count = folders.Count });
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
