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
    public class OutlookDiagnosticsController : ControllerBase
    {
        private readonly MailStore _mailStore;
        private readonly ChatStore _chatStore;
        private readonly CommandResultStore _commandResults;
        private readonly MockOutlookService _mockOutlook;
        private readonly IHubContext<NotificationHub> _hub;
        private readonly AddinStatusStore _addinStatus;

        public OutlookDiagnosticsController(
            MailStore mailStore,
            ChatStore chatStore,
            CommandResultStore commandResults,
            MockOutlookService mockOutlook,
            IHubContext<NotificationHub> hub,
            AddinStatusStore addinStatus)
        {
            _mailStore = mailStore;
            _chatStore = chatStore;
            _commandResults = commandResults;
            _mockOutlook = mockOutlook;
            _hub = hub;
            _addinStatus = addinStatus;
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
        /// Web UI、AI 或 MCP client 取得由 Hub cache 彙整的通訊錄關聯。
        /// </summary>
        [HttpGet("address-book")]
        public IActionResult GetAddressBook([FromQuery] string query = "", [FromQuery] int take = 200)
        {
            var contacts = _mailStore.GetAddressBookContacts(query, take);
            return Ok(new AddressBookResponse
            {
                Query = query ?? string.Empty,
                TotalCount = contacts.Count,
                Contacts = contacts,
            });
        }

        /// <summary>
        /// Web UI、AI 或 MCP client 依 email 查詢收件者是否出現在 Hub 已知互動裡。
        /// </summary>
        [HttpGet("address-book/lookup")]
        public IActionResult LookupAddressBookContact([FromQuery] string email)
        {
            if (string.IsNullOrWhiteSpace(email))
                return BadRequest(new { state = "failed", message = "email is required." });

            var contact = _mailStore.FindAddressBookContact(email);
            var suggestions = contact is null
                ? _mailStore.GetAddressBookContacts(email, 5)
                : new List<AddressBookContactDto>();

            return Ok(new AddressBookLookupResponse
            {
                Query = email,
                State = contact is null ? "unknown" : "known",
                Message = contact is null
                    ? "No cached mail or calendar relationship was found for this email."
                    : "Cached mail or calendar relationship found.",
                Contact = contact,
                Suggestions = suggestions,
            });
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
