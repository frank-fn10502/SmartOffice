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

        // ===================== Web UI 呼叫這些 endpoint =====================

        /// <summary>
        /// Web UI、AI 或 MCP client 要求 mails；Hub 會替 Outlook Add-in queue command。
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
        /// Web UI、AI 或 MCP client 要求 folder list。
        /// </summary>
        [HttpPost("request-folders")]
        public IActionResult RequestFolders()
        {
            _commandQueue.Enqueue(new PendingCommand { Type = "fetch_folders" });
            return Ok(new { status = "queued" });
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 要求 Outlook rule list。
        /// </summary>
        [HttpPost("request-rules")]
        public IActionResult RequestRules()
        {
            _commandQueue.Enqueue(new PendingCommand { Type = "fetch_rules" });
            return Ok(new { status = "queued" });
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 要求 Outlook master category list。
        /// </summary>
        [HttpPost("request-categories")]
        public IActionResult RequestCategories()
        {
            _commandQueue.Enqueue(new PendingCommand { Type = "fetch_categories" });
            return Ok(new { status = "queued" });
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 要求 Outlook calendar events。
        /// </summary>
        [HttpPost("request-calendar")]
        public IActionResult RequestCalendar([FromBody] FetchCalendarRequest? req)
        {
            var cmd = new PendingCommand
            {
                Type = "fetch_calendar",
                CalendarRequest = req ?? new FetchCalendarRequest()
            };
            _commandQueue.Enqueue(cmd);
            return Ok(new { commandId = cmd.Id, status = "queued" });
        }

        [HttpPost("request-mark-mail-read")]
        public IActionResult RequestMarkMailRead([FromBody] MailMarkerCommandRequest req)
        {
            return QueueMailMarkerCommand("mark_mail_read", req);
        }

        [HttpPost("request-mark-mail-unread")]
        public IActionResult RequestMarkMailUnread([FromBody] MailMarkerCommandRequest req)
        {
            return QueueMailMarkerCommand("mark_mail_unread", req);
        }

        [HttpPost("request-mark-mail-task")]
        public IActionResult RequestMarkMailTask([FromBody] MailMarkerCommandRequest req)
        {
            return QueueMailMarkerCommand("mark_mail_task", req);
        }

        [HttpPost("request-clear-mail-task")]
        public IActionResult RequestClearMailTask([FromBody] MailMarkerCommandRequest req)
        {
            return QueueMailMarkerCommand("clear_mail_task", req);
        }

        [HttpPost("request-set-mail-categories")]
        public IActionResult RequestSetMailCategories([FromBody] MailMarkerCommandRequest req)
        {
            return QueueMailMarkerCommand("set_mail_categories", req);
        }

        [HttpPost("request-update-mail-properties")]
        public IActionResult RequestUpdateMailProperties([FromBody] MailPropertiesCommandRequest req)
        {
            var cmd = new PendingCommand
            {
                Type = "update_mail_properties",
                MailPropertiesRequest = req
            };
            _commandQueue.Enqueue(cmd);
            return Ok(new { commandId = cmd.Id, status = "queued" });
        }

        [HttpPost("request-upsert-category")]
        public IActionResult RequestUpsertCategory([FromBody] CategoryCommandRequest req)
        {
            var cmd = new PendingCommand
            {
                Type = "upsert_category",
                CategoryRequest = req
            };
            _commandQueue.Enqueue(cmd);
            return Ok(new { commandId = cmd.Id, status = "queued" });
        }

        [HttpPost("request-create-folder")]
        public IActionResult RequestCreateFolder([FromBody] CreateFolderRequest req)
        {
            var cmd = new PendingCommand
            {
                Type = "create_folder",
                CreateFolderRequest = req
            };
            _commandQueue.Enqueue(cmd);
            return Ok(new { commandId = cmd.Id, status = "queued" });
        }

        [HttpPost("request-delete-folder")]
        public IActionResult RequestDeleteFolder([FromBody] DeleteFolderRequest req)
        {
            var cmd = new PendingCommand
            {
                Type = "delete_folder",
                DeleteFolderRequest = req
            };
            _commandQueue.Enqueue(cmd);
            return Ok(new { commandId = cmd.Id, status = "queued" });
        }

        [HttpPost("request-move-mail")]
        public IActionResult RequestMoveMail([FromBody] MoveMailRequest req)
        {
            var cmd = new PendingCommand
            {
                Type = "move_mail",
                MoveMailRequest = req
            };
            _commandQueue.Enqueue(cmd);
            return Ok(new { commandId = cmd.Id, status = "queued" });
        }

        private IActionResult QueueMailMarkerCommand(string type, MailMarkerCommandRequest req)
        {
            var cmd = new PendingCommand
            {
                Type = type,
                MailMarkerRequest = req
            };
            _commandQueue.Enqueue(cmd);
            return Ok(new { commandId = cmd.Id, status = "queued" });
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
            return Ok(_mailStore.GetFolders());
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

        // ===================== Outlook Add-in 呼叫這些 endpoint =====================

        /// <summary>
        /// Outlook Add-in 透過 long-poll 取得 pending command，timeout 為 30 秒。
        /// </summary>
        [HttpGet("poll")]
        public async Task<IActionResult> Poll(CancellationToken ct)
        {
            // 在受限網路內，Office 2016 Add-in 由 desktop process 主動發出
            // outbound HTTP，比要求 Hub 連入 desktop process 更容易部署。
            var cmd = await _commandQueue.DequeueAsync(TimeSpan.FromSeconds(30), ct);
            _addinStatus.RecordPoll(cmd?.Type);
            await _hub.Clients.All.SendAsync("AddinStatus", _addinStatus.GetStatus());
            if (cmd == null)
                return Ok(new { type = "none" });
            return Ok(cmd);
        }

        /// <summary>
        /// Outlook Add-in push mail results。
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
        /// Outlook Add-in push folder list。
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

        /// <summary>
        /// Outlook Add-in push rule list。
        /// </summary>
        [HttpPost("push-rules")]
        public async Task<IActionResult> PushRules([FromBody] List<OutlookRuleDto> rules)
        {
            _mailStore.SetRules(rules);
            _addinStatus.RecordPush("rules", rules.Count);
            await _hub.Clients.All.SendAsync("RulesUpdated", rules);
            await _hub.Clients.All.SendAsync("AddinStatus", _addinStatus.GetStatus());
            await _hub.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs());
            return Ok(new { count = rules.Count });
        }

        /// <summary>
        /// Outlook Add-in push master category list。
        /// </summary>
        [HttpPost("push-categories")]
        public async Task<IActionResult> PushCategories([FromBody] List<OutlookCategoryDto> categories)
        {
            _mailStore.SetCategories(categories);
            _addinStatus.RecordPush("categories", categories.Count);
            await _hub.Clients.All.SendAsync("CategoriesUpdated", categories);
            await _hub.Clients.All.SendAsync("AddinStatus", _addinStatus.GetStatus());
            await _hub.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs());
            return Ok(new { count = categories.Count });
        }

        /// <summary>
        /// Outlook Add-in push calendar events。
        /// </summary>
        [HttpPost("push-calendar")]
        public async Task<IActionResult> PushCalendar([FromBody] List<CalendarEventDto> events)
        {
            _mailStore.SetCalendarEvents(events);
            _addinStatus.RecordPush("calendar", events.Count);
            await _hub.Clients.All.SendAsync("CalendarUpdated", events);
            await _hub.Clients.All.SendAsync("AddinStatus", _addinStatus.GetStatus());
            await _hub.Clients.All.SendAsync("AddinLog", _addinStatus.GetLogs());
            return Ok(new { count = events.Count });
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
