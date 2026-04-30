using SmartOffice.Hub.Models;
using System.Collections.Concurrent;

namespace SmartOffice.Hub.Services
{
    public class MailStore
    {
        private readonly object _lock = new();
        private List<MailItemDto> _mails = new();
        private List<FolderDto> _folders = new();
        private List<OutlookRuleDto> _rules = new();
        private List<OutlookCategoryDto> _categories = new();
        private List<CalendarEventDto> _calendarEvents = new();

        public void SetMails(List<MailItemDto> mails)
        {
            lock (_lock) { _mails = new List<MailItemDto>(mails); }
        }

        public List<MailItemDto> GetMails()
        {
            lock (_lock) { return new List<MailItemDto>(_mails); }
        }

        public void SetFolders(List<FolderDto> folders)
        {
            lock (_lock) { _folders = new List<FolderDto>(folders); }
        }

        public List<FolderDto> GetFolders()
        {
            lock (_lock) { return new List<FolderDto>(_folders); }
        }

        public void SetRules(List<OutlookRuleDto> rules)
        {
            lock (_lock) { _rules = new List<OutlookRuleDto>(rules); }
        }

        public List<OutlookRuleDto> GetRules()
        {
            lock (_lock) { return new List<OutlookRuleDto>(_rules); }
        }

        public void SetCategories(List<OutlookCategoryDto> categories)
        {
            lock (_lock) { _categories = new List<OutlookCategoryDto>(categories); }
        }

        public List<OutlookCategoryDto> GetCategories()
        {
            lock (_lock) { return new List<OutlookCategoryDto>(_categories); }
        }

        public void SetCalendarEvents(List<CalendarEventDto> events)
        {
            lock (_lock) { _calendarEvents = new List<CalendarEventDto>(events); }
        }

        public List<CalendarEventDto> GetCalendarEvents()
        {
            lock (_lock) { return new List<CalendarEventDto>(_calendarEvents); }
        }
    }

    public class ChatStore
    {
        private readonly List<ChatMessageDto> _messages = new();
        private readonly object _lock = new();

        public void Add(ChatMessageDto msg)
        {
            lock (_lock) { _messages.Add(msg); }
        }

        public List<ChatMessageDto> GetAll()
        {
            lock (_lock) { return new List<ChatMessageDto>(_messages); }
        }
    }

    /// <summary>
    /// 儲存 Web UI、AI 或 MCP client 發出的 command，等待 Outlook Add-in 取走。
    /// </summary>
    public class CommandQueue
    {
        private readonly ConcurrentQueue<PendingCommand> _queue = new();
        private readonly SemaphoreSlim _signal = new(0);

        public void Enqueue(PendingCommand cmd)
        {
            _queue.Enqueue(cmd);
            _signal.Release();
        }

        /// <summary>
        /// Outlook Add-in 會 long-poll 這裡；timeout 時回傳 null。
        /// </summary>
        public async Task<PendingCommand?> DequeueAsync(TimeSpan timeout, CancellationToken ct = default)
        {
            if (await _signal.WaitAsync(timeout, ct))
            {
                _queue.TryDequeue(out var cmd);
                return cmd;
            }
            return null;
        }
    }

    public class AddinStatusStore
    {
        private readonly object _lock = new();
        private readonly List<AddinLogEntry> _logs = new();
        private bool _connected;
        private DateTime? _lastPollTime;
        private DateTime? _lastPushTime;
        private string _lastCommand = string.Empty;

        public void RecordPoll(string? commandType = null)
        {
            lock (_lock)
            {
                _connected = true;
                _lastPollTime = DateTime.Now;
                if (!string.IsNullOrEmpty(commandType))
                    _lastCommand = commandType;
                AddLogInternal("info", $"Add-in polled. Command dispatched: {commandType ?? "none"}");
            }
        }

        public void RecordPush(string pushType, int count)
        {
            lock (_lock)
            {
                _connected = true;
                _lastPushTime = DateTime.Now;
                AddLogInternal("info", $"Add-in pushed {pushType}: {count} items");
            }
        }

        public void AddLog(string level, string message)
        {
            lock (_lock) { AddLogInternal(level, message); }
        }

        private void AddLogInternal(string level, string message)
        {
            _logs.Add(new AddinLogEntry { Level = level, Message = message, Timestamp = DateTime.Now });
            if (_logs.Count > 500) _logs.RemoveAt(0);
        }

        public AddinStatusDto GetStatus()
        {
            lock (_lock)
            {
                // Add-in 每個 request cycle 都會 poll；目前 protocol 中，
                // 安靜但持續的 poll stream 就是最輕量的 heartbeat。
                if (_lastPollTime.HasValue && (DateTime.Now - _lastPollTime.Value).TotalSeconds > 90)
                    _connected = false;

                return new AddinStatusDto
                {
                    Connected = _connected,
                    LastPollTime = _lastPollTime,
                    LastPushTime = _lastPushTime,
                    LastCommand = _lastCommand
                };
            }
        }

        public List<AddinLogEntry> GetLogs()
        {
            lock (_lock) { return new List<AddinLogEntry>(_logs); }
        }
    }
}
