using SmartOffice.Hub.Models;
using System.Collections.Concurrent;

namespace SmartOffice.Hub.Services
{
    public class MailStore
    {
        private readonly object _lock = new();
        private List<MailItemDto> _mails = new();
        private List<FolderDto> _folders = new();

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
    /// Queue of commands from Web UI, AI, or MCP clients waiting for the Outlook Add-in.
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
        /// Outlook Add-in long-polls this. Returns command or null on timeout.
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
                // The add-in polls every request cycle; a quiet poll stream is the
                // best lightweight heartbeat available in the current protocol.
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
